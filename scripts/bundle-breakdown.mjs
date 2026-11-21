import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { gzipSync } from 'node:zlib';
import { build, version as esbuildVersion } from 'esbuild';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';
import { runLazyRuntime } from './bundle-runtime.mjs';

const output = process.argv[2];
const strictCsp = process.argv[3] === '--csp';
if (
  !output ||
  !output.endsWith('.json') ||
  (process.argv[3] && !strictCsp) ||
  process.argv.length > 4
)
  throw new Error('Usage: node scripts/bundle-breakdown.mjs NEW_REPORT.json [--csp]');
const root = process.cwd(),
  temp = mkdtempSync(join(tmpdir(), 'opensheet-bundles-'));
const hash = (data) => createHash('sha256').update(data).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  node: process.version,
  esbuild: esbuildVersion,
  scope:
    'Isolated pnpm-installed current workspace archives, esbuild browser minification and gzip; separate Worker, synthetic selected entry points and optional XLSX dynamic import. Localhost runtime checks use these lazy outputs; not a Vite, real-network transfer, runtime-memory, execution-time or complete consumer bundle measurement.',
  packages: [],
  scenarios: {},
};
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  mkdirSync(join(temp, 'archives'));
  const dependencies = { xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz' },
    overrides = {};
  if (strictCsp)
    dependencies.zod = JSON.parse(
      readFileSync(join(root, 'packages/core/package.json'), 'utf8'),
    ).dependencies.zod;
  for (const name of [
    'formula',
    'core',
    'renderer',
    'plugin-sdk',
    'opensheet',
    'adapter-sheetjs',
  ]) {
    const dir = join(root, 'packages', name),
      manifest = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
    const file = join(temp, 'archives', `${name}.tgz`);
    execFileSync('pnpm', ['pack', '--out', file], { cwd: dir, stdio: 'pipe', timeout: 120000 });
    dependencies[manifest.name] = overrides[manifest.name] = `file:./archives/${name}.tgz`;
    report.packages.push({
      name: manifest.name,
      version: manifest.version,
      sha256: hash(readFileSync(file)),
    });
  }
  writeFileSync(
    join(temp, 'package.json'),
    JSON.stringify({ private: true, type: 'module', dependencies, pnpm: { overrides } }),
  );
  execFileSync(
    'pnpm',
    ['install', '--prefer-offline', '--ignore-scripts', '--strict-peer-dependencies'],
    { cwd: temp, stdio: 'pipe', timeout: 120000 },
  );
  const url = 'https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz';
  const workspaceLock = readFileSync(join(root, 'pnpm-lock.yaml'), 'utf8');
  const escapedUrl = url.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const pinned = workspaceLock.match(
    new RegExp(`resolution: \{integrity: (sha512-[A-Za-z0-9+/=]+), tarball: ${escapedUrl}\}`),
  );
  assert.ok(pinned, 'Workspace SheetJS integrity is missing');
  const consumerLock = readFileSync(join(temp, 'pnpm-lock.yaml'), 'utf8');
  const original = `resolution: {tarball: ${url}}`;
  assert.equal(consumerLock.split(original).length, 2, 'Unexpected SheetJS lockfile shape');
  const lock = Buffer.from(
    consumerLock.replace(original, `resolution: {integrity: ${pinned[1]}, tarball: ${url}}`),
  );
  writeFileSync(join(temp, 'pnpm-lock.yaml'), lock);
  report.xlsxIntegrity = { source: 'pnpm-lock.yaml', integrity: pinned[1] };
  for (const name of ['formula', 'core', 'renderer', 'plugin-sdk', 'opensheet', 'adapter-sheetjs'])
    assert.ok(
      lock.toString('utf8').includes(`tarball: file:archives/${name}.tgz`),
      `${name}: archive missing from lockfile`,
    );
  rmSync(join(temp, 'node_modules'), { recursive: true, force: true });
  execFileSync(
    'pnpm',
    [
      'install',
      '--frozen-lockfile',
      '--prefer-offline',
      '--ignore-scripts',
      '--strict-peer-dependencies',
    ],
    { cwd: temp, stdio: 'pipe', timeout: 120000 },
  );
  assert.ok(
    lock.equals(readFileSync(join(temp, 'pnpm-lock.yaml'))),
    'Frozen reinstall changed the lockfile',
  );
  const lockPath = output.replace(/\.json$/, '') + '.consumer-lock.yaml';
  writeFileSync(lockPath, lock, { flag: 'wx' });
  report.lock = { path: lockPath, sha256: hash(lock), frozenReinstall: 'passed' };
  const plugins = [];
  if (strictCsp) {
    const shim = join(temp, 'csp-zod.js');
    const shimBytes = readFileSync(join(root, 'examples/framework-host/csp-zod.js'));
    writeFileSync(shim, shimBytes);
    report.cspShimSha256 = hash(shimBytes);
    plugins.push({
      name: 'zod-jitless',
      setup(builder) {
        builder.onResolve({ filter: /^zod$/ }, () => ({ path: shim }));
      },
    });
  }
  report.strictCsp = strictCsp;
  const entries = {
    'core.js':
      "import { createWorkbook, openWorkbook } from '@opensheetjs/core'; window.api = { createWorkbook, openWorkbook };",
    'editor.js':
      "import { createOpenSheet } from 'opensheet'; import 'opensheet/style.css'; window.api = { createOpenSheet };",
    'adapter.js':
      "import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs'; import * as XLSX from 'xlsx'; window.api = { fromSheetJS, toSheetJS, XLSX };",
    'lazy.js':
      "import { createOpenSheet } from 'opensheet'; import 'opensheet/style.css'; window.api = { createOpenSheet, loadXlsx: () => import('./xlsx-optional.js') };",
    'xlsx-optional.js':
      "import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs'; import * as XLSX from 'xlsx'; export { fromSheetJS, toSheetJS, XLSX };",
  };
  for (const [name, source] of Object.entries(entries)) writeFileSync(join(temp, name), source);
  report.entryHashes = Object.fromEntries(
    Object.entries(entries).map(([name, source]) => [name, hash(source)]),
  );
  const summarize = (meta) =>
    Object.fromEntries(
      Object.entries(meta.outputs).map(([file, info]) => {
        const contents = readFileSync(join(temp, file)),
          groups = {};
        for (const [input, contribution] of Object.entries(info.inputs)) {
          const match = input.match(
            /node_modules\/(?:\.pnpm\/[^/]+\/node_modules\/)?(@opensheetjs\/[^/]+|opensheet|xlsx|zod|immer|lucide)(?:\/|$)/,
          );
          const group = match?.[1] ?? (input.endsWith('.css') ? 'other-css' : 'entry-or-other');
          groups[group] = (groups[group] ?? 0) + contribution.bytesInOutput;
        }
        return [
          file,
          {
            bytes: contents.length,
            gzipBytes: gzipSync(contents, { level: 9 }).length,
            sha256: hash(contents),
            entryPoint: info.entryPoint ?? null,
            imports: info.imports,
            contributions: groups,
          },
        ];
      }),
    );
  for (const entry of ['core', 'editor', 'adapter']) {
    const result = await build({
      absWorkingDir: temp,
      entryPoints: [`${entry}.js`],
      bundle: true,
      minify: true,
      format: 'esm',
      platform: 'browser',
      target: 'es2022',
      outdir: `out/${entry}`,
      metafile: true,
      plugins,
      logLevel: 'silent',
    });
    report.scenarios[entry] = summarize(result.metafile);
  }
  const worker = await build({
    absWorkingDir: temp,
    entryPoints: ['node_modules/@opensheetjs/core/dist/engine.worker.js'],
    bundle: true,
    minify: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    outdir: 'out/worker',
    metafile: true,
    plugins,
    logLevel: 'silent',
  });
  report.scenarios.worker = summarize(worker.metafile);
  const lazy = await build({
    absWorkingDir: temp,
    entryPoints: ['lazy.js'],
    bundle: true,
    minify: true,
    splitting: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    outdir: 'out/lazy',
    metafile: true,
    plugins,
    logLevel: 'silent',
  });
  report.scenarios.lazy = summarize(lazy.metafile);
  const outputs = lazy.metafile.outputs,
    entry = Object.keys(outputs).find((name) => outputs[name].entryPoint === 'lazy.js');
  assert.ok(entry, 'Missing lazy entry');
  const initial = new Set(),
    visit = (file) => {
      if (initial.has(file)) return;
      initial.add(file);
      for (const dependency of outputs[file].imports)
        if (!dependency.external && dependency.kind !== 'dynamic-import') visit(dependency.path);
      if (outputs[file].cssBundle) visit(outputs[file].cssBundle);
    };
  visit(entry);
  const xlsxOutputs = Object.keys(report.scenarios.lazy).filter(
    (file) => report.scenarios.lazy[file].contributions.xlsx > 0,
  );
  assert.ok(xlsxOutputs.length, 'Optional XLSX code was not included in any output');
  assert.ok(
    xlsxOutputs.every((file) => !initial.has(file)),
    'XLSX code was included in the initial load',
  );
  report.lazyIsolation = { initial: [...initial], xlsxOutputs, status: 'passed' };
  report.runtime = await runLazyRuntime(temp, report.scenarios.lazy, xlsxOutputs, { strictCsp });
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  if (error.stdout) report.stdout = String(error.stdout).slice(-5000);
  if (error.stderr) report.stderr = String(error.stderr).slice(-5000);
  save();
  throw error;
} finally {
  rmSync(temp, { recursive: true, force: true });
}
