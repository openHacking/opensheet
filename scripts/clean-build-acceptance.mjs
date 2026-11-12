import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  cpSync,
  mkdtempSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative, sep } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/clean-build-acceptance.mjs NEW_REPORT.json');
const root = process.cwd();
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  pnpm: execFileSync('pnpm', ['--version'], { encoding: 'utf8' }).trim(),
  source: sourceFingerprint(),
  runs: [],
  scope:
    'Two fresh temporary workspace copies, frozen-lockfile prefer-offline installs, package builds and staged packs on one host using the same pnpm store and toolchain. Not independent hardware, a clean git commit, cross-platform provenance, registry publication or browser acceptance.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const temp = mkdtempSync(join(tmpdir(), 'opensheet-clean-build-'));
function sourceHash(base) {
  const files = [];
  function walk(path) {
    for (const item of readdirSync(path, { withFileTypes: true })) {
      if (item.name === 'node_modules' || item.name === 'dist') continue;
      const next = join(path, item.name);
      if (item.isDirectory()) walk(next);
      else if (item.isFile()) files.push(next);
    }
  }
  walk(join(base, 'packages'));
  for (const path of [
    'package.json',
    'pnpm-lock.yaml',
    'pnpm-workspace.yaml',
    'tsconfig.json',
    '.npmrc',
    'LICENSE',
    'scripts/build.mjs',
    'scripts/pack-candidate.mjs',
  ])
    files.push(join(base, path));
  const hash = createHash('sha256');
  for (const path of files.sort()) {
    const rel = relative(base, path),
      bytes = readFileSync(path);
    hash.update(`${rel.length}:${rel}:${bytes.length}:`);
    hash.update(bytes);
  }
  return { files: files.length, sha256: hash.digest('hex') };
}
try {
  for (let run = 1; run <= 2; run++) {
    const workspace = join(temp, `workspace-${run}`);
    mkdirSync(workspace);
    cpSync(join(root, 'packages'), join(workspace, 'packages'), {
      recursive: true,
      filter: (path) =>
        !relative(join(root, 'packages'), path)
          .split(sep)
          .some((part) => part === 'node_modules' || part === 'dist'),
    });
    mkdirSync(join(workspace, 'scripts'));
    for (const path of [
      'package.json',
      'pnpm-lock.yaml',
      'pnpm-workspace.yaml',
      'tsconfig.json',
      '.npmrc',
      'LICENSE',
      'scripts/build.mjs',
      'scripts/pack-candidate.mjs',
    ])
      cpSync(join(root, path), join(workspace, path));
    const data = { run, source: sourceHash(workspace), packages: [] };
    report.runs.push(data);
    save();
    console.log(`Workspace ${run}: installing frozen dependencies`);
    execFileSync('pnpm', ['install', '--frozen-lockfile', '--prefer-offline', '--ignore-scripts'], {
      cwd: workspace,
      encoding: 'utf8',
      stdio: 'pipe',
      timeout: 300000,
    });
    console.log(`Workspace ${run}: building packages`);
    execFileSync(process.execPath, ['scripts/build.mjs'], {
      cwd: workspace,
      encoding: 'utf8',
      stdio: 'pipe',
      timeout: 300000,
      maxBuffer: 8 * 1024 * 1024,
    });
    const names = readdirSync(join(workspace, 'packages'), { withFileTypes: true })
      .filter((d) => d.isDirectory())
      .map((d) => d.name)
      .sort();
    mkdirSync(join(workspace, 'archives'));
    for (const name of names) {
      const archive = join(workspace, 'archives', `${name}.tgz`);
      execFileSync(process.execPath, ['scripts/pack-candidate.mjs', name, archive], {
        cwd: workspace,
        encoding: 'utf8',
        stdio: 'pipe',
        timeout: 120000,
      });
      data.packages.push({ name, sha256: digest(readFileSync(archive)) });
    }
    save();
    console.log(`Workspace ${run}: built and packed ${names.length} packages`);
  }
  assert.deepEqual(report.runs[0].source, report.runs[1].source, 'workspace source copies differ');
  assert.deepEqual(
    report.runs[0].packages,
    report.runs[1].packages,
    'clean build package hashes differ',
  );
  report.identical = true;
  report.finishedAt = new Date().toISOString();
  save();
  console.log('Independent directories: hashes match');
} catch (error) {
  report.error = String(error);
  if (error.stdout) report.stdout = String(error.stdout).slice(-8000);
  if (error.stderr) report.stderr = String(error.stderr).slice(-8000);
  save();
  throw error;
} finally {
  rmSync(temp, { recursive: true, force: true });
}
