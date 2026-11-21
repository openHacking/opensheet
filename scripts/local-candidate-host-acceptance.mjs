import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { cpSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { chromium, firefox, webkit } from '@playwright/test';
import { preview } from 'vite';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
const artifacts = resolve(process.argv[3] ?? 'artifacts/opensheet-0.2.0-dev.0');
if (!output)
  throw new Error(
    'Usage: node scripts/local-candidate-host-acceptance.mjs NEW_REPORT.json [ARTIFACT_DIRECTORY]',
  );
const root = process.cwd(),
  temp = mkdtempSync(join(tmpdir(), 'opensheet-local-host-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  artifactDirectory: artifacts,
  results: [],
  scope:
    'Actual local candidate archives, isolated pnpm/Vite host under a subpath, three headless browsers. No registry publication of OpenSheet, strict CSP, SSR, native keyboard, memory or physical-device acceptance.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
let server;
try {
  save('wx');
  cpSync(join(root, 'examples/local-candidate'), temp, { recursive: true });
  report.example = Object.fromEntries(
    ['package.json', 'index.html', 'main.js', 'vite.config.js'].map((file) => [
      file,
      hash(readFileSync(join(temp, file))),
    ]),
  );
  mkdirSync(join(temp, 'archives'));
  const manifest = JSON.parse(readFileSync(join(artifacts, 'manifest.json'), 'utf8'));
  report.archives = [];
  for (const folder of ['opensheet', 'core', 'formula', 'plugin-sdk', 'renderer']) {
    const name = `${folder}.tgz`,
      source = join(artifacts, name),
      bytes = readFileSync(source);
    assert.equal(
      hash(bytes),
      manifest.packages.find((item) => item.file === name)?.sha256,
      `${name}: checksum mismatch`,
    );
    writeFileSync(join(temp, 'archives', name), bytes);
    report.archives.push({ name, sha256: hash(bytes) });
  }
  execFileSync(
    'pnpm',
    ['install', '--prefer-offline', '--ignore-scripts', '--strict-peer-dependencies'],
    { cwd: temp, encoding: 'utf8', stdio: 'pipe', timeout: 120000 },
  );
  const lock = readFileSync(join(temp, 'pnpm-lock.yaml')),
    lockPath = output.replace(/\.json$/, '') + '.consumer-lock.yaml';
  for (const folder of ['opensheet', 'core', 'formula', 'plugin-sdk', 'renderer'])
    assert.ok(
      lock.toString('utf8').includes(`tarball: file:archives/${folder}.tgz`),
      `${folder}: lockfile does not resolve local archive`,
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
    { cwd: temp, encoding: 'utf8', stdio: 'pipe', timeout: 120000 },
  );
  assert.ok(
    lock.equals(readFileSync(join(temp, 'pnpm-lock.yaml'))),
    'Frozen reinstall changed the lockfile',
  );
  writeFileSync(lockPath, lock, { flag: 'wx' });
  report.lock = { path: lockPath, sha256: hash(lock), frozenReinstall: 'passed' };
  execFileSync('pnpm', ['run', 'build'], {
    cwd: temp,
    encoding: 'utf8',
    stdio: 'pipe',
    timeout: 120000,
  });
  server = await preview({
    root: temp,
    configFile: join(temp, 'vite.config.js'),
    preview: { host: '127.0.0.1', port: 0, strictPort: true },
  });
  const address = server.httpServer.address();
  const url = `http://127.0.0.1:${address.port}/embedded/sheets/`;
  for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
    const browser = await launcher.launch();
    try {
      const page = await browser.newPage(),
        errors = [],
        workerUrls = [];
      page.on('pageerror', (error) => errors.push(String(error)));
      page.on('request', (request) => {
        if (request.url().includes('engine.worker'))
          workerUrls.push(new URL(request.url()).pathname);
      });
      await page.goto(url);
      await page.waitForFunction(
        () => window.candidateExample || document.body.dataset.error,
        null,
        { timeout: 30000 },
      );
      const initial = await page.evaluate(async () => {
        if (document.body.dataset.error) throw Error(document.body.dataset.error);
        const { app, book } = window.candidateExample,
          sheet = book.getSheets()[0],
          element = document.querySelector('.opensheet');
        app.selection.set({
          sheetId: sheet.id,
          startRow: 1,
          endRow: 2,
          startColumn: 0,
          endColumn: 1,
        });
        return {
          values: await sheet.range('A1:C1').getValues(),
          height: element.getBoundingClientRect().height,
          stylesheets: document.querySelectorAll('link[rel=stylesheet]').length,
        };
      });
      const grid = page.getByRole('grid', { exact: true });
      await grid.focus();
      await grid.press('Enter');
      const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
      await editor.fill('42');
      await editor.press('Enter');
      await page.waitForFunction(
        async () => {
          try {
            return (
              (await window.candidateExample.book.getSheets()[0].range('A2').getValues())[0][0] ===
              42
            );
          } catch (error) {
            if (error?.code === 'REVISION_CONFLICT') return false;
            throw error;
          }
        },
        null,
        { timeout: 30000 },
      );
      const attachmentRace = await page.evaluate(async () => {
        const { app, book, backupWorkbook } = window.candidateExample;
        const snapshot = JSON.parse(await backupWorkbook(book));
        const target = new book.constructor({
          ...snapshot,
          workbookId: `candidate_target_${crypto.randomUUID()}`,
        });
        await target.ready();
        const backupValues = await target.getSheets()[0].range('A1:C2').getValues();
        const ready = target.ready.bind(target);
        let release;
        const pending = new Promise((resolve) => (release = resolve));
        target.ready = async () => {
          await pending;
          return ready();
        };
        const result = app.attachWorkbook(target).then(
          () => 'attached',
          (error) => error.code,
        );
        await target.deleteStorage();
        release();
        const code = await result;
        return {
          code,
          backupValues,
          same: app.getWorkbook() === book,
          value: await book.getSheets()[0].range('A2').getValues(),
          grids: document.querySelectorAll('.os-grid').length,
        };
      });
      const edited = await page.evaluate(async () => {
        const { app, book, openWorkbook, backupWorkbook, restoreBackup, createOpenSheet } =
          window.candidateExample;
        const backup = await backupWorkbook(book);
        const value = (await book.getSheets()[0].range('A2').getValues())[0][0];
        const id = book.id;
        app.dispose();
        const removed = !document.querySelector('.opensheet');
        const restoreApp = createOpenSheet({ container: '#sheet' });
        const restored = await restoreBackup(restoreApp, backup);
        const restoredValues = await restored.getSheets()[0].range('A1:C2').getValues();
        await restored.deleteStorage();
        restoreApp.dispose();
        const reopened = await openWorkbook(id);
        try {
          return {
            value,
            removed,
            restoredValues,
            reopened: await reopened.getSheets()[0].range('A1:C2').getValues(),
          };
        } finally {
          await reopened.close();
        }
      });
      const result = {
        engine,
        browser: browser.version(),
        initial,
        attachmentRace,
        edited,
        workerUrls,
        errors,
      };
      report.results.push(result);
      save();
      assert.deepEqual(initial.values, [[7, 5, 12]]);
      assert.equal(initial.height, 500);
      assert.equal(initial.stylesheets, 1);
      assert.deepEqual(attachmentRace, {
        code: 'DISPOSED',
        backupValues: [
          [7, 5, 12],
          [42, null, null],
        ],
        same: true,
        value: [[42]],
        grids: 1,
      });
      assert.deepEqual(edited, {
        value: 42,
        removed: true,
        restoredValues: [
          [7, 5, 12],
          [42, null, null],
        ],
        reopened: [
          [7, 5, 12],
          [42, null, null],
        ],
      });
      assert.ok(
        workerUrls.filter((path) => path.startsWith('/embedded/sheets/assets/')).length >= 2,
      );
      assert.deepEqual(errors, []);
      console.log(`${engine}: passed`);
    } finally {
      await browser.close();
    }
  }
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  if (error.stdout) report.stdout = String(error.stdout).slice(-5000);
  if (error.stderr) report.stderr = String(error.stderr).slice(-5000);
  save();
  throw error;
} finally {
  if (server) await new Promise((done) => server.httpServer.close(done));
  rmSync(temp, { recursive: true, force: true });
}
