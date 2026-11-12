import { execFileSync } from 'node:child_process';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { copyFileSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
const artifactDirectory = process.argv[3] ? resolve(process.argv[3]) : undefined;
const artifactManifest = artifactDirectory
  ? JSON.parse(readFileSync(join(artifactDirectory, 'manifest.json'), 'utf8'))
  : undefined;
if (!output) throw new Error('Usage: node scripts/esbuild-consumer-acceptance.mjs NEW_REPORT');
const root = process.cwd();
const consumer = mkdtempSync(join(tmpdir(), 'opensheet-esbuild-'));
const report = {
  source: sourceFingerprint(),
  artifactDirectory: artifactDirectory ?? null,
  startedAt: new Date().toISOString(),
  esbuild: (await import('esbuild')).version,
  scope:
    'Packed editor and dependencies; separate esbuild browser, CSS and Worker bundles, strict CSP with Zod jitless shim, localhost subpath and real IndexedDB/Web Locks. No SSR framework, other bundlers or physical devices.',
  packages: [],
  results: [],
};
const hash = (data) => createHash('sha256').update(data).digest('hex');
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
let server;
try {
  mkdirSync(join(consumer, 'tarballs'));
  const dependencies = { zod: '4.6.6' },
    overrides = {};
  for (const folder of ['formula', 'core', 'renderer', 'plugin-sdk', 'opensheet']) {
    const packageDir = join(root, 'packages', folder);
    const manifest = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'));
    const archive = join(consumer, 'tarballs', `${folder}.tgz`);
    if (artifactManifest) {
      copyFileSync(join(artifactDirectory, `${folder}.tgz`), archive);
      assert.equal(
        hash(readFileSync(archive)),
        artifactManifest.packages.find((item) => item.name === manifest.name)?.sha256,
        `${folder}: candidate archive mismatch`,
      );
    } else execFileSync('pnpm', ['pack', '--out', archive], { cwd: packageDir, stdio: 'pipe' });
    dependencies[manifest.name] = overrides[manifest.name] = `file:./tarballs/${folder}.tgz`;
    report.packages.push({
      name: manifest.name,
      version: manifest.version,
      sha256: hash(readFileSync(archive)),
    });
  }
  writeFileSync(
    join(consumer, 'package.json'),
    JSON.stringify({ private: true, type: 'module', dependencies, pnpm: { overrides } }),
  );
  execFileSync(
    'pnpm',
    ['install', '--prefer-offline', '--ignore-scripts', '--strict-peer-dependencies'],
    {
      cwd: consumer,
      stdio: 'pipe',
      timeout: 120000,
    },
  );
  const lock = readFileSync(join(consumer, 'pnpm-lock.yaml'));
  const lockPath = output.replace(/\.json$/, '') + '.consumer-lock.yaml';
  writeFileSync(lockPath, lock, { flag: 'wx' });
  report.lock = { path: lockPath, sha256: hash(lock) };
  const entry = `
import { OpenSheet, openWorkbook } from 'opensheet';
import 'opensheet/style.css';
try {
  const editor = new OpenSheet({container: '#root'});
  const book = await editor.createWorkbook({sheets:[{name:'Data',rows:4,columns:2}]});
  const id = book.id, sheet = book.getSheets()[0];
  await sheet.range('A1:B1').setValues([[6,7]]);
  await sheet.range('A2').setFormulas([['A1+B1']]);
  const before = await sheet.range('A1:B2').getValues();
  await editor.ready();
  const node = document.querySelector('.opensheet');
  const ui = { grid: !!document.querySelector('.os-grid'), display: getComputedStyle(node).display, height: node.getBoundingClientRect().height, stylesheets: document.querySelectorAll('link[rel=stylesheet]').length };
  editor.selection.set({sheetId:sheet.id,startRow:2,endRow:3,startColumn:0,endColumn:1});
  window.readEditedCell = async () => (await sheet.range('A3').getValues())[0][0];
  window.editorReady = true;
  await new Promise(resolve => { window.finishEditor = resolve; });
  ui.edited = await window.readEditedCell();
  editor.dispose();
  ui.removed = !document.querySelector('.opensheet');
  const reopened = await openWorkbook(id);
  const after = await reopened.getSheets()[0].range('A1:B2').getValues();
  ui.reopenedEdit = (await reopened.getSheets()[0].range('A3').getValues())[0][0];
  reopened.dispose();
  document.body.dataset.result = JSON.stringify({before,after,ui});
} catch (error) { document.body.dataset.error = String(error && (error.stack || error)); }
`;
  writeFileSync(join(consumer, 'main.js'), entry);
  const shim = join(consumer, 'csp-zod.js');
  writeFileSync(shim, readFileSync(join(root, 'examples', 'framework-host', 'csp-zod.js')));
  const zodShim = {
    name: 'zod-jitless',
    setup(builder) {
      builder.onResolve({ filter: /^zod$/ }, () => ({ path: shim }));
    },
  };
  mkdirSync(join(consumer, 'out'));
  const bundle = join(consumer, 'out', 'main.js');
  const built = await build({
    entryPoints: [join(consumer, 'main.js')],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    outfile: bundle,
    metafile: true,
    plugins: [zodShim],
  });
  const worker = join(consumer, 'out', 'engine.worker.js');
  await build({
    entryPoints: [
      join(consumer, 'node_modules', '@opensheetjs', 'core', 'dist', 'engine.worker.js'),
    ],
    bundle: true,
    format: 'esm',
    platform: 'browser',
    target: 'es2022',
    outfile: worker,
    plugins: [zodShim],
    banner: {
      js: `self.addEventListener('securitypolicyviolation', e => self.postMessage({__cspViolation: {directive:e.effectiveDirective, blocked:e.blockedURI}}));`,
    },
  });
  writeFileSync(
    join(consumer, 'out', 'index.html'),
    '<!doctype html><html><head><link rel="stylesheet" href="./main.css"></head><body><main id="root" class="os-framework-host"></main><script type="module" src="./main.js"></script></body></html>',
  );
  writeFileSync(
    join(consumer, 'out', 'probe.worker.js'),
    `
self.addEventListener('securitypolicyviolation', e => self.postMessage({__cspViolation:{directive:e.effectiveDirective,blocked:e.blockedURI}}));
self.onmessage = () => { try { new Function('return 1')(); self.postMessage({blocked:false}); } catch { self.postMessage({blocked:true}); } };
`,
  );
  writeFileSync(
    join(consumer, 'out', 'probe.js'),
    `try { new Function('return 1')(); document.body.dataset.probeBlocked = 'false'; } catch { document.body.dataset.probeBlocked = 'true'; }`,
  );
  report.assets = {
    bundleBytes: readFileSync(bundle).length,
    bundleSha256: hash(readFileSync(bundle)),
    workerBytes: readFileSync(worker).length,
    workerSha256: hash(readFileSync(worker)),
    cssBytes: readFileSync(join(consumer, 'out', 'main.css')).length,
    cssSha256: hash(readFileSync(join(consumer, 'out', 'main.css'))),
    inputs: Object.keys(built.metafile.inputs).length,
  };
  const csp =
    "default-src 'none'; script-src 'self'; worker-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; font-src 'self'; base-uri 'none'; object-src 'none'";
  report.csp = csp;
  report.shimSha256 = hash(readFileSync(shim));
  const requests = [];
  server = createServer((req, res) => {
    requests.push(req.url);
    const file =
      req.url === '/embedded/sheets/'
        ? 'index.html'
        : req.url === '/embedded/sheets/main.js'
          ? 'main.js'
          : req.url === '/embedded/sheets/main.css'
            ? 'main.css'
            : req.url === '/embedded/sheets/engine.worker.js'
              ? 'engine.worker.js'
              : req.url === '/embedded/sheets/probe.worker.js'
                ? 'probe.worker.js'
                : req.url === '/embedded/sheets/probe.js'
                  ? 'probe.js'
                  : undefined;
    if (!file) {
      res.writeHead(404);
      res.end();
      return;
    }
    res.writeHead(200, {
      'Content-Type': file.endsWith('.js')
        ? 'text/javascript'
        : file.endsWith('.css')
          ? 'text/css'
          : 'text/html',
      'Content-Security-Policy': csp,
    });
    res.end(readFileSync(join(consumer, 'out', file)));
  });
  await new Promise((resolveListen) => server.listen(0, '127.0.0.1', resolveListen));
  const url = `http://127.0.0.1:${server.address().port}/embedded/sheets/`;
  for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
    const browser = await launcher.launch();
    try {
      const page = await browser.newPage();
      const errors = [];
      page.on('pageerror', (error) => errors.push(String(error)));
      await page.addInitScript(() => {
        window.cspEvents = [];
        document.addEventListener('securitypolicyviolation', (event) =>
          window.cspEvents.push({
            source: 'document',
            directive: event.effectiveDirective,
            blocked: event.blockedURI,
          }),
        );
        const OriginalWorker = window.Worker;
        window.Worker = class extends OriginalWorker {
          constructor(...args) {
            super(...args);
            this.addEventListener('message', ({ data }) => {
              if (data?.__cspViolation)
                window.cspEvents.push({ source: 'worker', ...data.__cspViolation });
            });
          }
        };
      });
      const start = requests.length;
      await page.goto(url);
      await page.waitForFunction(() => window.editorReady || document.body.dataset.error, null, {
        timeout: 30000,
      });
      const initError = await page.evaluate(() => document.body.dataset.error);
      if (initError) throw new Error(`Editor initialization: ${initError}`);
      const grid = page.getByRole('grid', { exact: true });
      await grid.focus();
      await grid.press('Enter');
      const cellEditor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
      await cellEditor.fill('42');
      await cellEditor.press('Enter');
      await page.waitForFunction(
        async () => {
          try {
            return (await window.readEditedCell()) === 42;
          } catch (error) {
            if (error?.code === 'REVISION_CONFLICT') return false;
            throw error;
          }
        },
        null,
        { timeout: 30000 },
      );
      await page.evaluate(() => window.finishEditor());
      await page.waitForFunction(
        () => document.body.dataset.result || document.body.dataset.error,
        null,
        { timeout: 30000 },
      );
      const outcome = await page.evaluate(() => ({
        result: document.body.dataset.result && JSON.parse(document.body.dataset.result),
        error: document.body.dataset.error,
        cspEvents: window.cspEvents,
      }));
      const expected = JSON.stringify([
        [6, 7],
        [13, null],
      ]);
      const workerRequests = requests
        .slice(start)
        .filter((path) => path === '/embedded/sheets/engine.worker.js').length;
      if (
        outcome.error ||
        outcome.cspEvents.length ||
        errors.length ||
        JSON.stringify(outcome.result?.before) !== expected ||
        JSON.stringify(outcome.result?.after) !== expected ||
        !outcome.result?.ui.grid ||
        outcome.result?.ui.display !== 'flex' ||
        outcome.result?.ui.height !== 500 ||
        outcome.result?.ui.stylesheets !== 1 ||
        !outcome.result?.ui.removed ||
        outcome.result?.ui.edited !== 42 ||
        outcome.result?.ui.reopenedEdit !== 42 ||
        workerRequests < 2
      ) {
        report.results.push({ engine, outcome, errors, workerRequests, passed: false });
        save();
        throw new Error(`${engine}: ${JSON.stringify({ outcome, errors, workerRequests })}`);
      }
      const probe = await page.evaluate(async () => {
        const documentBlocked = await new Promise((resolveProbe, rejectProbe) => {
          const script = document.createElement('script');
          script.src = './probe.js';
          script.onload = () => resolveProbe(document.body.dataset.probeBlocked === 'true');
          script.onerror = () => rejectProbe(new Error('CSP probe script failed'));
          document.head.append(script);
        });
        const workerBlocked = await new Promise((resolveProbe, rejectProbe) => {
          const worker = new Worker('./probe.worker.js', { type: 'module' });
          let blocked,
            violation = false;
          const timeout = setTimeout(() => {
            worker.terminate();
            if (blocked === undefined) rejectProbe(new Error('CSP probe timed out'));
            else resolveProbe({ blocked, violation });
          }, 1000);
          worker.onerror = () => {
            worker.terminate();
            rejectProbe(new Error('CSP probe worker failed'));
          };
          worker.onmessage = ({ data }) => {
            if (typeof data.blocked === 'boolean') blocked = data.blocked;
            if (data.__cspViolation) violation = true;
            if (blocked !== undefined && violation) {
              clearTimeout(timeout);
              worker.terminate();
              resolveProbe({ blocked, violation });
            }
          };
          worker.postMessage('run');
        });
        return { documentBlocked, workerBlocked };
      });
      await page.waitForFunction(
        () => window.cspEvents.some((e) => e.source === 'document'),
        null,
        { timeout: 3000 },
      );
      const probeEvents = await page.evaluate(() => window.cspEvents);
      if (!probe.documentBlocked || !probe.workerBlocked.blocked)
        throw new Error(`${engine}: CSP negative control failed: ${JSON.stringify(probe)}`);
      report.results.push({
        engine,
        before: outcome.result.before,
        after: outcome.result.after,
        ui: outcome.result.ui,
        workerRequests,
        pageErrors: errors,
        cspEvents: outcome.cspEvents,
        negativeControl: { ...probe, events: probeEvents },
      });
      save();
      console.log(`${engine}: passed`);
    } catch (error) {
      if (!report.results.some((item) => item.engine === engine)) {
        report.results.push({ engine, passed: false, error: String(error) });
        save();
      }
      throw error;
    } finally {
      await browser.close();
    }
  }
} finally {
  if (server) await new Promise((resolveClose) => server.close(resolveClose));
  rmSync(consumer, { recursive: true, force: true });
}
