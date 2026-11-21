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
const exampleDirectory = resolve(process.argv[4] ?? 'examples/local-candidate');
if (!output)
  throw new Error(
    'Usage: node scripts/local-candidate-host-acceptance.mjs NEW_REPORT.json [ARTIFACT_DIRECTORY] [EXAMPLE_DIRECTORY]',
  );
const root = process.cwd(),
  temp = mkdtempSync(join(tmpdir(), 'opensheet-local-host-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  artifactDirectory: artifacts,
  exampleDirectory,
  results: [],
  scope:
    'Actual local candidate archives, isolated pnpm/Vite host under a subpath, three headless browsers. No registry publication of OpenSheet, strict CSP, SSR, native keyboard, memory or physical-device acceptance.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
let server;
try {
  save('wx');
  cpSync(exampleDirectory, temp, { recursive: true });
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
      let originalId = await page.evaluate(() => window.candidateExample.book.id);
      await page.reload();
      await page.waitForFunction(
        () => window.candidateExample || document.body.dataset.error,
        null,
        { timeout: 30000 },
      );
      const refreshed = await page.evaluate(async () => {
        if (document.body.dataset.error) throw Error(document.body.dataset.error);
        const book = window.candidateExample.book;
        return {
          id: book.id,
          values: await book.getSheets()[0].range('A1:C2').getValues(),
          readOnly: book.isReadOnly,
        };
      });
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
      const prematureDownloads = [];
      page.on('download', (download) => prematureDownloads.push(download.suggestedFilename()));
      await page.evaluate(() => {
        const { app, book } = window.candidateExample;
        app.selection.set({
          sheetId: book.getSheets()[0].id,
          startRow: 2,
          endRow: 3,
          startColumn: 0,
          endColumn: 1,
        });
      });
      await page.getByRole('grid', { exact: true }).focus();
      await page.getByRole('grid', { exact: true }).press('Enter');
      const draftEditor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
      await draftEditor.fill('123');
      await page.waitForFunction(() => window.candidateExample.app.hasDraft);
      const draftOutcome = Promise.race([
        page.waitForEvent('download', { timeout: 5000 }).then(
          () => 'downloaded',
          () => 'timeout',
        ),
        page
          .waitForFunction(
            () =>
              document
                .querySelector('#status')
                .textContent.includes('Finish editing and wait for saving'),
            null,
            { timeout: 5000 },
          )
          .then(
            () => 'blocked',
            () => 'timeout',
          ),
      ]);
      await page.evaluate(() => document.querySelector('#backup').click());
      assert.equal(await draftOutcome, 'blocked', 'Draft backup should not start a download');
      const draftGuard = await page.evaluate(() => ({
        draft: window.candidateExample.app.hasDraft,
        enabled: !document.querySelector('#backup').disabled,
      }));
      assert.deepEqual(draftGuard, { draft: true, enabled: true });
      await draftEditor.press('Escape');
      await page.evaluate(() => {
        const book = window.candidateExample.book,
          worker = book.worker;
        const post = worker.postMessage.bind(worker),
          handler = worker.onmessage;
        let id;
        worker.postMessage = function (message, ...rest) {
          if (message.method === 'execute') id = message.id;
          return post(message, ...rest);
        };
        worker.onmessage = (event) => {
          if (event.data.id === id) window.heldBackupResponse = event;
          else handler.call(worker, event);
        };
        window.releasePendingBackupWrite = async () => {
          worker.postMessage = post;
          worker.onmessage = handler;
          handler.call(worker, window.heldBackupResponse);
          await window.pendingBackupWrite;
          await book.undo();
        };
        window.pendingBackupWrite = book
          .getSheets()[0]
          .range('A3')
          .setValues([[777]]);
      });
      await page.waitForFunction(() => window.heldBackupResponse);
      await page.locator('#backup').click();
      await page.waitForFunction(
        () =>
          document
            .querySelector('#status')
            .textContent.includes('Finish editing and wait for saving'),
        null,
        { timeout: 5000 },
      );
      const pendingGuard = await page.evaluate(
        () => window.candidateExample.book.persistenceState.pending,
      );
      assert.ok(pendingGuard > 0);
      await page.evaluate(() => window.releasePendingBackupWrite());
      await page.evaluate(() => {
        const book = window.candidateExample.book,
          original = book.toJSON;
        let release;
        const gate = new Promise((resolve) => (release = resolve));
        book.toJSON = async function (...args) {
          const snapshot = await original.apply(this, args);
          window.backupSnapshotHeld = true;
          await gate;
          return snapshot;
        };
        window.releaseHeldBackup = () => {
          book.toJSON = original;
          release();
        };
      });
      await page.locator('#backup').click();
      await page.waitForFunction(() => window.backupSnapshotHeld);
      await page.evaluate(async () =>
        window.candidateExample.book
          .getSheets()[0]
          .range('A3')
          .setValues([[888]]),
      );
      await page.evaluate(() => window.releaseHeldBackup());
      await page.waitForFunction(
        () =>
          document
            .querySelector('#status')
            .textContent.includes('Workbook changed while creating the backup'),
        null,
        { timeout: 5000 },
      );
      const changedGuard = await page.evaluate(async () => {
        const book = window.candidateExample.book;
        const value = await book.getSheets()[0].range('A3').getValues();
        await book.undo();
        return { value, enabled: !document.querySelector('#backup').disabled };
      });
      assert.deepEqual(changedGuard, { value: [[888]], enabled: true });
      await page.evaluate(() => {
        const book = window.candidateExample.book,
          original = book.toJSON;
        window.backupSnapshotHeld = false;
        let release;
        const gate = new Promise((resolve) => (release = resolve));
        book.toJSON = async function (...args) {
          const snapshot = await original.apply(this, args);
          window.backupSnapshotHeld = true;
          await gate;
          return snapshot;
        };
        window.releaseHeldBackup = () => {
          book.toJSON = original;
          release();
        };
      });
      await page.locator('#backup').click();
      await page.waitForFunction(() => window.backupSnapshotHeld);
      await page.evaluate(() => {
        const { app, book } = window.candidateExample;
        app.selection.set({
          sheetId: book.getSheets()[0].id,
          startRow: 2,
          endRow: 3,
          startColumn: 0,
          endColumn: 1,
        });
      });
      await page.getByRole('grid', { exact: true }).focus();
      await page.getByRole('grid', { exact: true }).press('Enter');
      await draftEditor.fill('321');
      await page.waitForFunction(() => window.candidateExample.app.hasDraft);
      await page.evaluate(() => window.releaseHeldBackup());
      await page.waitForFunction(
        () =>
          document
            .querySelector('#status')
            .textContent.includes('Workbook changed while creating the backup'),
        null,
        { timeout: 5000 },
      );
      const lateDraftGuard = {
        draft: await draftEditor.inputValue(),
        active: await page.evaluate(() => window.candidateExample.app.hasDraft),
      };
      assert.deepEqual(lateDraftGuard, { draft: '321', active: true });
      await draftEditor.press('Escape');
      assert.deepEqual(prematureDownloads, []);
      const guardedDownloads = [...prematureDownloads];
      await page.evaluate(() => {
        window.originalBlob = globalThis.Blob;
        globalThis.Blob = class extends window.originalBlob {
          get size() {
            return 64 * 1024 * 1024 + 1;
          }
        };
      });
      await page.locator('#backup').click();
      await page.waitForFunction(() =>
        document
          .querySelector('#status')
          .textContent.includes('Backup exceeds the 64 MiB example limit'),
      );
      await page.evaluate(() => {
        globalThis.Blob = window.originalBlob;
      });
      await page.evaluate(() => {
        Object.defineProperty(File.prototype, 'size', {
          configurable: true,
          get: () => 64 * 1024 * 1024 + 1,
        });
      });
      await page.locator('#restore').setInputFiles({
        name: 'large.json',
        mimeType: 'application/json',
        buffer: Buffer.from('{}'),
      });
      await page.waitForFunction(
        () =>
          document.querySelector('#status').textContent.startsWith('Could not restore backup;') &&
          document.querySelector('#status').textContent.includes('64 MiB'),
      );
      await page.evaluate(() => {
        delete File.prototype.size;
      });
      const limitResult = await page.evaluate(async () => ({
        id: window.candidateExample.book.id,
        savedId: localStorage.getItem('opensheet-local-candidate-workbook-id'),
        value: await window.candidateExample.book.getSheets()[0].range('A2').getValues(),
        enabled:
          !document.querySelector('#backup').disabled &&
          !document.querySelector('#restore').disabled,
      }));
      assert.deepEqual(limitResult, {
        id: originalId,
        savedId: originalId,
        value: [[42]],
        enabled: true,
      });
      const [download] = await Promise.all([
        page.waitForEvent('download'),
        page.locator('#backup').click(),
      ]);
      assert.equal(download.suggestedFilename(), 'workbook.opensheet.json');
      const backupBytes = readFileSync(await download.path());
      await page.locator('#restore').setInputFiles({
        name: 'invalid.json',
        mimeType: 'application/json',
        buffer: Buffer.from('{'),
      });
      await page.waitForFunction(() =>
        document.querySelector('#status').textContent.startsWith('Could not restore backup;'),
      );
      const badRestore = await page.evaluate(async () => ({
        id: window.candidateExample.book.id,
        savedId: localStorage.getItem('opensheet-local-candidate-workbook-id'),
        value: await window.candidateExample.book.getSheets()[0].range('A2').getValues(),
      }));
      await page.locator('#restore').setInputFiles({
        name: 'invalid-schema.json',
        mimeType: 'application/json',
        buffer: Buffer.from('{}'),
      });
      await page.waitForFunction(() =>
        document.querySelector('#status').textContent.startsWith('Could not restore backup;'),
      );
      const badSchema = await page.evaluate(async () => ({
        id: window.candidateExample.book.id,
        savedId: localStorage.getItem('opensheet-local-candidate-workbook-id'),
        value: await window.candidateExample.book.getSheets()[0].range('A2').getValues(),
      }));
      assert.deepEqual(badSchema, badRestore);
      await page.evaluate(async () =>
        window.candidateExample.book
          .getSheets()[0]
          .range('A2')
          .setValues([[99]]),
      );
      await page
        .locator('#restore')
        .setInputFiles({ name: 'backup.json', mimeType: 'application/json', buffer: backupBytes });
      await page.waitForFunction(
        () => document.querySelector('#status').textContent === 'Backup restored.',
      );
      const restoredViaUI = await page.evaluate(async () => ({
        id: window.candidateExample.book.id,
        savedId: localStorage.getItem('opensheet-local-candidate-workbook-id'),
        values: await window.candidateExample.book.getSheets()[0].range('A1:C2').getValues(),
      }));
      assert.notEqual(restoredViaUI.id, originalId);
      assert.equal(restoredViaUI.savedId, restoredViaUI.id);
      assert.deepEqual(badRestore, { id: originalId, savedId: originalId, value: [[42]] });
      assert.deepEqual(restoredViaUI.values, [
        [7, 5, 12],
        [42, null, null],
      ]);
      originalId = restoredViaUI.id;
      await page.evaluate(() => {
        const prototype = Object.getPrototypeOf(localStorage);
        window.originalStorageSetItem = prototype.setItem;
        prototype.setItem = function (key, value) {
          if (key === 'opensheet-local-candidate-workbook-id')
            throw new Error('Injected storage failure');
          return window.originalStorageSetItem.call(this, key, value);
        };
      });
      await page
        .locator('#restore')
        .setInputFiles({ name: 'backup.json', mimeType: 'application/json', buffer: backupBytes });
      await page.waitForFunction(() =>
        document
          .querySelector('#status')
          .textContent.startsWith('Workbook restored, but its ID could not be saved.'),
      );
      const saveFailure = await page.evaluate(async () => {
        const book = window.candidateExample.book;
        return {
          id: book.id,
          savedId: localStorage.getItem('opensheet-local-candidate-workbook-id'),
          values: await book.getSheets()[0].range('A1:C2').getValues(),
        };
      });
      assert.notEqual(saveFailure.id, originalId);
      assert.equal(saveFailure.savedId, originalId);
      assert.deepEqual(saveFailure.values, [
        [7, 5, 12],
        [42, null, null],
      ]);
      await page.evaluate((id) => {
        Object.getPrototypeOf(localStorage).setItem = window.originalStorageSetItem;
        localStorage.setItem('opensheet-local-candidate-workbook-id', id);
      }, saveFailure.id);
      originalId = saveFailure.id;
      await page.reload();
      await page.waitForFunction(
        () => window.candidateExample || document.body.dataset.error,
        null,
        { timeout: 30000 },
      );
      const restoredAfterReload = await page.evaluate(async () => {
        if (document.body.dataset.error) throw Error(document.body.dataset.error);
        const book = window.candidateExample.book;
        return {
          id: book.id,
          values: await book.getSheets()[0].range('A1:C2').getValues(),
          readOnly: book.isReadOnly,
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
      const missingId = `missing_${Date.now()}`;
      await page.evaluate(
        (id) => localStorage.setItem('opensheet-local-candidate-workbook-id', id),
        missingId,
      );
      await page.reload();
      await page.waitForFunction(
        () => window.candidateExample || document.body.dataset.error,
        null,
        { timeout: 30000 },
      );
      const failure = await page.evaluate(() => ({
        retainedId: localStorage.getItem('opensheet-local-candidate-workbook-id'),
        status: document.querySelector('#status').textContent,
        opened: !!window.candidateExample,
      }));
      await page.evaluate(
        (id) => localStorage.setItem('opensheet-local-candidate-workbook-id', id),
        originalId,
      );
      await page.reload();
      await page.waitForFunction(
        () => window.candidateExample || document.body.dataset.error,
        null,
        { timeout: 30000 },
      );
      const recovered = await page.evaluate(async () => {
        if (document.body.dataset.error) throw Error(document.body.dataset.error);
        const book = window.candidateExample.book;
        return {
          id: book.id,
          values: await book.getSheets()[0].range('A1:C2').getValues(),
          readOnly: book.isReadOnly,
        };
      });
      const result = {
        engine,
        browser: browser.version(),
        initial,
        refreshed,
        limitResult,
        badRestore,
        badSchema,
        restoredViaUI,
        saveFailure,
        restoredAfterReload,
        failure,
        recovered,
        attachmentRace,
        draftGuard,
        pendingGuard,
        changedGuard,
        lateDraftGuard,
        guardedDownloads,
        edited,
        workerUrls,
        errors,
      };
      report.results.push(result);
      save();
      assert.deepEqual(initial.values, [[7, 5, 12]]);
      assert.equal(initial.height, 500);
      assert.equal(initial.stylesheets, 1);
      const expectedReload = {
        id: originalId,
        values: [
          [7, 5, 12],
          [42, null, null],
        ],
        readOnly: false,
      };
      assert.deepEqual(refreshed, { ...expectedReload, id: badRestore.id });
      assert.deepEqual(restoredAfterReload, expectedReload);
      assert.equal(failure.retainedId, missingId);
      assert.equal(failure.opened, false);
      assert.match(failure.status, /^Could not open workbook:/);
      assert.deepEqual(recovered, expectedReload);
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
