import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { createServer } from 'node:http';
import { copyFileSync, mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
const args = process.argv.slice(3);
const injectHeadAbort = args.includes('--head-abort');
const artifactOptions = args.filter((arg) => arg.startsWith('--artifacts='));
if (
  artifactOptions.length > 1 ||
  args.filter((arg) => arg === '--head-abort').length > 1 ||
  args.some((arg) => arg !== '--head-abort' && !arg.startsWith('--artifacts='))
)
  throw new Error('Unknown or duplicate option');
const artifacts = resolve(
  artifactOptions[0]?.slice('--artifacts='.length) ?? 'artifacts/opensheet-0.2.0-dev.0',
);
if (!output)
  throw new Error('Usage: node scripts/historical-upgrade-acceptance.mjs NEW_REPORT.json');
const temp = mkdtempSync(join(tmpdir(), 'opensheet-upgrade-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  node: process.version,
  results: [],
  injectHeadAbort,
  artifactDirectory: artifacts,
  scope:
    'Published npm core 0.1.1 and locally packed 0.2.0-dev.0, separately installed and esbuild-bundled in one localhost origin. Optional head-transaction abort is injected into the served candidate Worker. Synthetic workbook, one sample per browser; not complete historical compatibility, physical crash testing, independent toolchain, physical devices or safe downgrade certification.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
let server;
try {
  save('wx');
  const archives = join(temp, 'archives');
  mkdirSync(archives);
  const packed = JSON.parse(
    execFileSync(
      'npm',
      ['pack', '@opensheetjs/core@0.1.1', '--pack-destination', archives, '--json'],
      { cwd: temp, encoding: 'utf8', timeout: 120000 },
    ),
  )[0];
  const oldArchive = join(archives, packed.filename);
  report.historical = {
    name: packed.name,
    version: packed.version,
    integrity: packed.integrity,
    sha256: hash(readFileSync(oldArchive)),
  };
  for (const folder of ['core', 'formula'])
    copyFileSync(join(artifacts, `${folder}.tgz`), join(archives, `${folder}.tgz`));
  report.candidate = {
    coreSha256: hash(readFileSync(join(archives, 'core.tgz'))),
    formulaSha256: hash(readFileSync(join(archives, 'formula.tgz'))),
  };
  const candidateManifest = JSON.parse(readFileSync(join(artifacts, 'manifest.json'), 'utf8'));
  for (const folder of ['core', 'formula']) {
    const expected = candidateManifest.packages.find((item) => item.file === `${folder}.tgz`);
    assert.equal(
      hash(readFileSync(join(archives, `${folder}.tgz`))),
      expected?.sha256,
      `${folder}: candidate archive mismatch`,
    );
  }
  for (const version of ['old', 'new']) {
    const dir = join(temp, version);
    mkdirSync(dir);
    const dependencies =
      version === 'old'
        ? {
            '@opensheetjs/core': `file:../archives/${packed.filename}`,
            '@opensheetjs/formula': '0.1.1',
          }
        : {
            '@opensheetjs/core': 'file:../archives/core.tgz',
            '@opensheetjs/formula': 'file:../archives/formula.tgz',
          };
    const overrides =
      version === 'old'
        ? {}
        : {
            '@opensheetjs/core': 'file:../archives/core.tgz',
            '@opensheetjs/formula': 'file:../archives/formula.tgz',
          };
    writeFileSync(
      join(dir, 'package.json'),
      JSON.stringify({ private: true, type: 'module', dependencies, pnpm: { overrides } }),
    );
    execFileSync(
      'pnpm',
      ['install', '--prefer-offline', '--ignore-scripts', '--strict-peer-dependencies'],
      { cwd: dir, encoding: 'utf8', stdio: 'pipe', timeout: 120000 },
    );
    const dist = join(dir, 'node_modules/@opensheetjs/core/dist');
    for (const [input, name] of [
      ['index.js', 'main.js'],
      ['engine.worker.js', 'engine.worker.js'],
    ])
      await build({
        entryPoints: [join(dist, input)],
        bundle: true,
        platform: 'browser',
        format: 'esm',
        target: 'es2022',
        outfile: join(dir, name),
      });
    const lock = readFileSync(join(dir, 'pnpm-lock.yaml'));
    const lockPath = output.replace(/\.json$/, '') + `.${version}-consumer-lock.yaml`;
    writeFileSync(lockPath, lock, { flag: 'wx' });
    report[`${version}Assets`] = {
      main: hash(readFileSync(join(dir, 'main.js'))),
      worker: hash(readFileSync(join(dir, 'engine.worker.js'))),
      lock: { path: lockPath, sha256: hash(lock) },
    };
  }
  const injection = `
let abortNextHead = false;
addEventListener('message', event => { if (event.data.armHeadAbort) { abortNextHead = true; postMessage({headFaultArmed:true}); } });
const originalHeadPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
 if (abortNextHead && this.name === 'heads') { abortNextHead=false; this.transaction.abort(); postMessage({headFaultTriggered:true}); throw new DOMException('Injected head transaction abort','AbortError'); }
 return originalHeadPut.apply(this,args);
};`;
  const candidateWorker = readFileSync(join(temp, 'new', 'engine.worker.js'));
  const servedCandidateWorker = injectHeadAbort
    ? Buffer.concat([candidateWorker, Buffer.from(injection)])
    : candidateWorker;
  report.newAssets.servedWorker = hash(servedCandidateWorker);
  server = createServer((req, res) => {
    const path = req.url;
    if (path === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end('<!doctype html><body>Upgrade test</body>');
      return;
    }
    const match = /^\/(old|new)\/(main|engine\.worker)\.js$/.exec(path);
    if (!match) {
      res.writeHead(404);
      res.end();
      return;
    }
    res.writeHead(200, { 'Content-Type': 'text/javascript' });
    const bytes = readFileSync(join(temp, match[1], `${match[2]}.js`));
    res.end(
      injectHeadAbort && match[1] === 'new' && match[2] === 'engine.worker'
        ? servedCandidateWorker
        : bytes,
    );
  });
  await new Promise((done) => server.listen(0, '127.0.0.1', done));
  const url = `http://127.0.0.1:${server.address().port}/`;
  for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
    const browser = await launcher.launch();
    try {
      const page = await browser.newPage();
      await page.goto(url);
      const result = await page.evaluate(
        async ({ injectHeadAbort }) => {
          const old = await import('/old/main.js'),
            next = await import('/new/main.js');
          const database = 'upgrade-' + crypto.randomUUID();
          const book = await old.createWorkbook(
            {
              sheets: [
                { name: 'Data', rows: 4, columns: 3 },
                { name: 'Rates', rows: 4, columns: 3 },
              ],
            },
            { database },
          );
          const id = book.id,
            [sheet, rates] = book.getSheets();
          await sheet.range('A1:B1').setValues([[7, 5]]);
          await sheet.range('C1').setFormulas([['A1+B1']]);
          await sheet.range('B1').setStyle({ bold: true, numberFormat: '0.00' });
          await rates.range('A1').setValues([[3]]);
          await rates.range('B1').setFormulas([['Data!A1+3']]);
          await sheet.range('B2').setValues([['merged']]);
          await sheet.range('B2:C2').merge();
          await sheet.setRowHeight(2, 37);
          await sheet.setRowHidden(3, true);
          await sheet.setColumnWidth(1, 143);
          await sheet.setFreeze(1, 1);
          await sheet.range('A1').setValues([[8]]);
          if (!(await book.undo())) throw Error('old undo unavailable');
          const oldValues = await sheet.range('A1:C1').getValues();
          const metadata = async (workbook) => {
            const sheets = workbook.getSheets(),
              data = workbook.sheetData(sheets[0].id);
            return {
              names: sheets.map((entry) => entry.name),
              rates: await sheets[1].range('A1:B1').getValues(),
              rateFormula: (await sheets[1].range('B1').getFormulas())[0][0],
              mainFormula: (await sheets[0].range('C1').getFormulas())[0][0],
              mergedValue: (await sheets[0].range('B2').getValues())[0][0],
              merges: data.merges,
              freeze: data.freeze,
              rowSize: data.rows[data.rowOrder[2]]?.size,
              rowHidden: data.rows[data.rowOrder[3]]?.hidden,
              columnSize: data.columns[data.columnOrder[1]]?.size,
            };
          };
          const oldMetadata = await metadata(book);
          const file = await book.toJSON();
          const parts = [];
          for await (const part of book.streamBinary()) parts.push(part);
          const historicalBinary = new Blob(parts);
          await book.close();
          const snapshot = async () => {
            const db = await new Promise((resolveDb, reject) => {
              const r = indexedDB.open(database);
              r.onsuccess = () => resolveDb(r.result);
              r.onerror = () => reject(r.error);
            });
            try {
              const stores = [...db.objectStoreNames].sort(),
                out = {};
              for (const name of stores) {
                const tx = db.transaction(name, 'readonly'),
                  store = tx.objectStore(name);
                const keys = await new Promise((resolveKeys, reject) => {
                  const q = store.getAllKeys();
                  q.onsuccess = () => resolveKeys(q.result);
                  q.onerror = () => reject(q.error);
                });
                const values = await new Promise((resolveValues, reject) => {
                  const q = store.getAll();
                  q.onsuccess = () => resolveValues(q.result);
                  q.onerror = () => reject(q.error);
                });
                out[name] = keys.map((key, i) => [
                  key,
                  values[i] instanceof Uint8Array ? Array.from(values[i]) : values[i],
                ]);
              }
              return JSON.stringify(out);
            } finally {
              db.close();
            }
          };
          let json;
          try {
            const reader = new next.FileReader(file);
            try {
              const range = reader.getSheetById(file.sheets[0].id).range('A1:C1');
              json = { values: range.getValues(), display: range.getDisplayValues() };
            } finally {
              reader.dispose();
            }
          } catch (error) {
            json = { error: { code: error.code, message: error.message } };
          }
          const imports = {};
          for (const format of ['json', 'binary']) {
            let target = await next.createWorkbook(
              { sheets: [{ name: 'Target', rows: 4, columns: 3 }] },
              { database },
            );
            try {
              let targetSheet = target.getSheets()[0];
              await targetSheet.range('A1').setValues([['keep']]);
              const revision = target.revision;
              let invalid;
              if (format === 'json') invalid = { ...file, schemaVersion: 99 };
              else {
                invalid = new Uint8Array(await historicalBinary.arrayBuffer());
                invalid[3] = 99;
              }
              let error;
              try {
                if (format === 'json') await target.importJSON(invalid);
                else await target.importBinary(invalid);
              } catch (failure) {
                error = failure.code;
              }
              const rejected = {
                error,
                revisionUnchanged: target.revision === revision,
                value: (await targetSheet.range('A1').getValues())[0][0],
              };
              const rejectedId = target.id;
              await target.close();
              target = await next.openWorkbook(rejectedId, { database });
              targetSheet = target.getSheets()[0];
              rejected.reopenedRevisionUnchanged = target.revision === revision;
              rejected.reopenedValue = (await targetSheet.range('A1').getValues())[0][0];
              rejected.undo = await target.undo();
              rejected.redo = await target.redo();
              rejected.afterRedo = (await targetSheet.range('A1').getValues())[0][0];
              if (format === 'json') await target.importJSON(file);
              else await target.importBinary(historicalBinary);
              let active = target.getSheets()[0];
              const values = await active.range('A1:C1').getValues(),
                display = await target.display(active.id, 0, 1);
              const importedId = target.id;
              await target.close();
              target = await next.openWorkbook(importedId, { database });
              active = target.getSheets()[0];
              imports[format] = {
                rejected,
                values,
                display,
                reopened: await active.range('A1:C1').getValues(),
                metadata: await metadata(target),
              };
            } finally {
              await target.deleteStorage();
              target.dispose();
            }
          }
          const before = await snapshot();
          let fault;
          if (injectHeadAbort) {
            const candidate = await next.openWorkbook(id, { database });
            const revision = candidate.revision;
            const worker = candidate.worker;
            let triggered = false;
            worker.addEventListener('message', (event) => {
              if (event.data.headFaultTriggered) triggered = true;
            });
            await new Promise((resolveArmed) => {
              worker.addEventListener('message', function handler(event) {
                if (event.data.headFaultArmed) {
                  worker.removeEventListener('message', handler);
                  resolveArmed();
                }
              });
              worker.postMessage({ method: 'cancelRequest', args: [-1], armHeadAbort: true });
            });
            let error;
            try {
              await candidate
                .getSheets()[0]
                .range('A1')
                .setValues([[99]]);
            } catch (failure) {
              error = failure.code;
            }
            const revisionUnchanged = candidate.revision === revision;
            await candidate.close();
            const unchanged = (await snapshot()) === before;
            const historical = await old.openWorkbook(id, { database });
            let values;
            try {
              values = await historical.getSheets()[0].range('A1:C1').getValues();
            } finally {
              await historical.close();
            }
            fault = { triggered, error, revisionUnchanged, unchanged, historicalValues: values };
          }
          let upgrade;
          try {
            const opened = await next.openWorkbook(id, { database });
            try {
              const active = opened.getSheets()[0],
                values = await active.range('A1:C1').getValues(),
                redo = await opened.redo(),
                redone = await active.range('A1:C1').getValues(),
                redoneRates = await opened.getSheets()[1].range('A1:B1').getValues(),
                undo = await opened.undo(),
                restored = await active.range('A1:C1').getValues();
              const cell = await opened.getCell(active.id, 0, 1);
              upgrade = {
                values,
                redo,
                redone,
                redoneRates,
                undo,
                restored,
                style: opened.getStyle(cell?.styleId),
                display: await opened.display(active.id, 0, 1),
              };
            } finally {
              await opened.close();
            }
            const reopened = await next.openWorkbook(id, { database });
            try {
              upgrade.reopened = await reopened.getSheets()[0].range('A1:C1').getValues();
              upgrade.metadata = await metadata(reopened);
            } finally {
              await reopened.close();
            }
          } catch (error) {
            upgrade = {
              error: { code: error.code, message: error.message },
              unchanged: (await snapshot()) === before,
            };
          }
          let oldReopen;
          try {
            const reopened = await old.openWorkbook(id, { database });
            try {
              oldReopen = {
                values: await reopened.getSheets()[0].range('A1:C1').getValues(),
                metadata: await metadata(reopened),
              };
            } finally {
              await reopened.close();
            }
          } catch (error) {
            oldReopen = { error: { code: error.code, message: error.message } };
          }
          indexedDB.deleteDatabase(database);
          return {
            oldValues,
            oldMetadata,
            schemaVersion: file.schemaVersion,
            json,
            imports,
            fault,
            upgrade,
            oldReopen,
          };
        },
        { injectHeadAbort },
      );
      report.results.push({ engine, browser: browser.version(), ...result });
      save();
      const expectedMetadata = {
        names: ['Data', 'Rates'],
        rates: [[3, 10]],
        rateFormula: 'Data!A1+3',
        mainFormula: 'A1+B1',
        mergedValue: 'merged',
        merges: [{ startRow: 1, endRow: 2, startColumn: 1, endColumn: 3 }],
        freeze: { rows: 1, columns: 1 },
        rowSize: 37,
        rowHidden: true,
        columnSize: 143,
      };
      assert.deepEqual(result.oldValues, [[7, 5, 12]], `${engine}: historical values`);
      assert.deepEqual(result.oldMetadata, expectedMetadata, `${engine}: historical metadata`);
      assert.equal(result.schemaVersion, 3, `${engine}: historical schema`);
      assert.deepEqual(
        result.json,
        { values: [[7, 5, 12]], display: [['7', '5.00', '12']] },
        `${engine}: JSON reader`,
      );
      for (const format of ['json', 'binary']) {
        assert.deepEqual(
          result.imports[format],
          {
            rejected: {
              error: 'UNSUPPORTED_VERSION',
              revisionUnchanged: true,
              value: 'keep',
              reopenedRevisionUnchanged: true,
              reopenedValue: 'keep',
              undo: true,
              redo: true,
              afterRedo: 'keep',
            },
            values: [[7, 5, 12]],
            display: '5.00',
            reopened: [[7, 5, 12]],
            metadata: expectedMetadata,
          },
          `${engine}: ${format} import and rollback`,
        );
      }
      assert.deepEqual(result.upgrade?.values, [[7, 5, 12]], `${engine}: upgrade values`);
      assert.equal(result.upgrade?.redo, true, `${engine}: redo`);
      assert.deepEqual(result.upgrade?.redone, [[8, 5, 13]], `${engine}: redone values`);
      assert.deepEqual(
        result.upgrade?.redoneRates,
        [[3, 11]],
        `${engine}: redone cross-sheet formula`,
      );
      assert.equal(result.upgrade?.undo, true, `${engine}: undo`);
      assert.deepEqual(result.upgrade?.restored, [[7, 5, 12]], `${engine}: undo values`);
      assert.deepEqual(result.upgrade?.reopened, [[7, 5, 12]], `${engine}: reopened values`);
      assert.deepEqual(result.upgrade?.metadata, expectedMetadata, `${engine}: upgraded metadata`);
      assert.equal(result.upgrade?.style?.bold, true, `${engine}: bold style`);
      assert.equal(result.upgrade?.display, '5.00', `${engine}: formatted value`);
      assert.deepEqual(result.oldReopen?.values, [[7, 5, 12]], `${engine}: historical reopen`);
      assert.deepEqual(
        result.oldReopen?.metadata,
        expectedMetadata,
        `${engine}: historical metadata reopen`,
      );
      if (injectHeadAbort)
        assert.deepEqual(
          result.fault,
          {
            triggered: true,
            error: 'AbortError',
            revisionUnchanged: true,
            unchanged: true,
            historicalValues: [[7, 5, 12]],
          },
          `${engine}: interrupted publication`,
        );
      console.log(`${engine}: ${JSON.stringify(result)}`);
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
  if (server) await new Promise((done) => server.close(done));
  rmSync(temp, { recursive: true, force: true });
}
