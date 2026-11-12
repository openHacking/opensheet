import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from '../../scripts/benchmark-fingerprint.mjs';

const core = fileURLToPath(new URL('../../packages/core/src/', import.meta.url));

async function installHarness(page: Page) {
  await page.route('**/validation-backup-harness', (route) =>
    route.fulfill({
      contentType: 'text/html',
      body: '<!doctype html><title>Migration harness</title>',
    }),
  );
  await page.route('**/validation-backup-test.worker.js*', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
import { Engine } from ${JSON.stringify(`/@fs${core}engine.ts`)};
import { Database } from ${JSON.stringify(`/@fs${core}storage.ts`)};
import { createSnapshot } from ${JSON.stringify(`/@fs${core}model.ts`)};
import { encodeWorkbookFile } from ${JSON.stringify(`/@fs${core}file-codec.ts`)};
import { encodeBlock } from ${JSON.stringify(`/@fs${core}block-codec.ts`)};
import { writeBinary } from ${JSON.stringify(`/@fs${core}binary-file.ts`)};
import { upgradeValidationFile } from ${JSON.stringify(`/@fs${core}validation-file-migration.ts`)};
const params = new URL(self.location.href).searchParams;
const phase = params.get('phase');
const database = params.get('database') ?? 'validation-backup-' + crypto.randomUUID();
if (phase === 'benchmark') {
  const rows = 10000, columns = 100;
  const binary = params.get('format') === 'binary';
  const snapshot = createSnapshot({sheets:[{name:'Dense',rows,columns}]});
  const id = snapshot.workbookId;
  const file = encodeWorkbookFile(snapshot);
  const text = binary ? '' : JSON.stringify(upgradeValidationFile(file));
  const split = text.indexOf('"blocks":[]') + '"blocks":['.length;
  const parts = [text.slice(0, split)];
  let pages = 0;
  async function* fixtures() {
  for (let row = 0; row < rows; row += 64) for (let column = 0; column < columns; column += 32) {
    const positions = [], values = [];
    const cells = {};
    for (let r = 0; r < 64 && row + r < rows; r++) for (let c = 0; c < 32 && column + c < columns; c++) {
      const rowId = row + r, columnId = column + c, value = rowId * columns + columnId;
      if (binary) cells[rowId+':'+columnId] = {rowId,columnId,input:{type:'number',value}};
      else { positions.push(r * 32 + c); values.push(value); }
    }
    pages++;
    yield binary ? {sheet:0,row:row/64,column:column/32,bytes:encodeBlock(cells,[],[],{})}
      : {row:row/64,column:column/32,data:{numbers:{positions,values}}};
  }
  }
  if (binary) {
    parts.length = 0;
    for await (const chunk of writeBinary(file, fixtures())) parts.push(chunk);
  } else {
    for await (const block of fixtures()) {
      if (pages > 1) parts.push(',');
      parts.push(JSON.stringify(block));
    }
    parts.push(text.slice(split));
  }
  const backup = new Blob(parts);
  const phases = {};
  let db = await Database.open(database);
  let engine = new Engine(db,id,undefined,(event) => {
    if (typeof event.elapsedMs === 'number') phases[event.phase] = event.elapsedMs;
  });
  await engine.open(snapshot);
  const start = performance.now();
  const result = await (binary ? engine.importBinary(backup) : engine.importRuleFreeValidationBackup(backup));
  const importMs = performance.now() - start;
  engine.close();
  const reopenStart = performance.now();
  db = await Database.open(database);
  engine = new Engine(db,id);
  const reopened = await engine.open();
  const reopenMs = performance.now() - reopenStart;
  const sheetId = reopened.snapshot.sheets[0].id;
  const values = [];
  for (const [row,column] of [[0,0],[64,32],[rows-1,columns-1]]) {
    const cell = await engine.read(sheetId,{startRow:row,endRow:row+1,startColumn:column,endColumn:column+1});
    values.push(cell.calculated[row+':'+column]);
  }
  const verifyStart = performance.now();
  let verifiedCells = 0;
  for (let startRow = 0; startRow < rows; startRow += 64) {
    const endRow = Math.min(rows, startRow + 64);
    const page = await engine.read(sheetId,{startRow,endRow,startColumn:0,endColumn:columns});
    for (let row = startRow; row < endRow; row++) for (let column = 0; column < columns; column++) {
      if (page.calculated[row+':'+column] !== row * columns + column)
        throw new Error('Restored cell mismatch at '+row+':'+column);
      verifiedCells++;
    }
  }
  const get = (request) => new Promise((resolve,reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  const transaction = db.db.transaction(['heads','records','sizes']);
  const keys = IDBKeyRange.bound(id + '/', id + '/\uffff');
  const [head,recordKeys,records,sizeKeys,sizes] = await Promise.all([
    get(transaction.objectStore('heads').get(id)),
    get(transaction.objectStore('records').getAllKeys(keys)),
    get(transaction.objectStore('records').getAll(keys)),
    get(transaction.objectStore('sizes').getAllKeys(keys)),
    get(transaction.objectStore('sizes').getAll(keys)),
  ]);
  const decoder = new TextDecoder();
  const ledger = sizes.map(entry => JSON.parse(decoder.decode(entry.subarray(1))));
  const expectedLedger = recordKeys.map((key,index) => [key,records[index].byteLength]);
  const physicalBytes = head.byteLength + [...records,...sizes].reduce((sum,entry) => sum + entry.byteLength,0);
  const verifiedLedger = JSON.stringify(recordKeys) === JSON.stringify(sizeKeys) &&
    JSON.stringify(ledger) === JSON.stringify(expectedLedger) && physicalBytes === reopened.bytes;
  const verifyMs = performance.now() - verifyStart;
  const scratch = await db.hasAny('scratch', IDBKeyRange.bound(id + '/', id + '/\\uffff'));
  engine.close();
  postMessage({format:binary?'binary':'validation-json',cells:rows*columns,pages,backupBytes:backup.size,storedBytes:reopened.bytes,importMs,reopenMs,phases,
    revision:reopened.revision,verifiedBytes:result.bytes === reopened.bytes,verifiedCells,verifiedLedger,verifyMs,values,scratch});
} else if (phase === 'metadata') {
  const snapshot = createSnapshot({sheets:[{name:'Data',rows:130,columns:5},{name:'Other',rows:2,columns:2}]});
  const id = snapshot.workbookId;
  const sheet = snapshot.sheets[0];
  snapshot.dateSystem = '1904';
  snapshot.extensions['fixture'] = { provenance: 'mixed restore' };
  snapshot.styles.custom = {bold:true,numberFormat:'0.00'};
  sheet.rows['0'] = {size:42,hidden:true};
  sheet.columns['1'] = {size:140};
  sheet.freeze = {rows:1,columns:1};
  sheet.merges = [{startRow:1,endRow:2,startColumn:3,endColumn:5}];
  sheet.cells['0:0'] = {rowId:0,columnId:0,input:{type:'number',value:12},styleId:'custom',note:'source',link:{target:'https://example.com/'}};
  sheet.cells['64:2'] = {rowId:64,columnId:2,input:{type:'formula',expression:'A1*2'}};
  snapshot.sheets[1].hidden = true;
  const file = upgradeValidationFile(encodeWorkbookFile(snapshot));
  const empty = createSnapshot({sheets:[{name:'Original',rows:2,columns:2}]});
  empty.workbookId = id;
  let db = await Database.open(database);
  let engine = new Engine(db,id);
  await engine.open(empty);
  await engine.importRuleFreeValidationBackup(new Blob([JSON.stringify(file)]));
  engine.close();
  db = await Database.open(database);
  engine = new Engine(db,id);
  const view = await engine.open();
  const restored = view.snapshot.sheets[0];
  const first = await engine.cell(restored.id,0,0);
  const formula = await engine.cell(restored.id,64,2);
  const calculated = await engine.read(restored.id,{startRow:64,endRow:65,startColumn:2,endColumn:3});
  const result = {dateSystem:view.snapshot.dateSystem,extensions:view.snapshot.extensions,
    names:view.snapshot.sheets.map(sheet => [sheet.name,sheet.hidden]),
    rows:restored.rows,columns:restored.columns,freeze:restored.freeze,merges:restored.merges,
    first,style:view.snapshot.styles[first.styleId],formula,value:calculated.calculated['64:2']};
  engine.close();
  postMessage(result);
} else if (phase === 'reopen') {
  const id = params.get('id');
  const db = await Database.open(database);
  const engine = new Engine(db,id);
  const view = await engine.open();
  const cells = await engine.read(view.snapshot.sheets[0].id,{startRow:0,endRow:65,startColumn:0,endColumn:1});
  const scratch = await db.hasAny('scratch', IDBKeyRange.bound(id + '/', id + '/\\uffff'));
  engine.close();
  postMessage({revision:view.revision, values:[cells.calculated['0:0'] ?? null, cells.calculated['64:0'] ?? null], scratch});
} else {
const snapshot = createSnapshot({sheets: [{name:'Data', rows:80, columns:2}]});
const id = snapshot.workbookId;
const scope = IDBKeyRange.bound(id + '/', id + '/\\uffff');
let db = await Database.open(database);
let engine = new Engine(db, id);
await engine.open(snapshot);
const file = upgradeValidationFile(encodeWorkbookFile(snapshot));
file.sheets[0].blocks.push({row:0,column:0,data:{numbers:{positions:[0],values:[42]}}});
file.sheets[0].blocks.push({row:1,column:0,data:{numbers:{positions:[0],values:[43]}}});
const source = () => new Blob([JSON.stringify(file)]);
if (phase === 'scratch' || phase === 'heads') {
  const original = IDBObjectStore.prototype.put;
  let armed = true;
  IDBObjectStore.prototype.put = function(...args) {
    const request = original.apply(this,args);
    if (armed && this.name === phase) {
      armed = false;
      // Do not resume the importer after this durable transaction completes.
      this.transaction.oncomplete = () => postMessage({phase, database, id});
    }
    return request;
  };
  await engine.importRuleFreeValidationBackup(source());
  throw new Error('Import unexpectedly acknowledged after interception');
}

const base = await db.get('heads', id);
const baseline = JSON.stringify(base);
const status = async () => ({unchanged: JSON.stringify(await db.get('heads', id)) === baseline,
  scratch: await db.hasAny('scratch', scope), revision: engine.view().revision, bytes: engine.view().bytes});
const run = async () => {
  try { await engine.importRuleFreeValidationBackup(source()); return 'success'; }
  catch (error) { return error.code ?? error.name; }
};
const before = await status();
let sent = false, cancelReachedScratch = false, sourceCancelled = false;
const stalled = new ReadableStream({
  async pull(controller) {
    if (!sent) {
      sent = true;
      controller.enqueue(new TextEncoder().encode(JSON.stringify(file).slice(0,-1)));
    } else {
      cancelReachedScratch = await db.hasAny('scratch',scope);
      engine.cancel();
    }
  },
  cancel() { sourceCancelled = true; },
}, {highWaterMark:0});
let cancelled;
try { await engine.importRuleFreeValidationBackup(stalled); cancelled = 'success'; }
catch (error) { cancelled = error.code ?? error.name; }
const afterCancel = await status();
engine.begin();
file.sheets[0].validation = [{id:'required',range:{startRow:0,endRow:1,startColumn:0,endColumn:1},rule:{kind:'list',values:[42],allowBlank:false,policy:'reject'}}];
const rejected = await run();
const afterRules = await status();
file.sheets[0].validation = [];
let armed = true, injected = false;
const original = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = original.apply(this,args);
  if (this.name === 'heads' && armed) { armed = false; injected = true; this.transaction.abort(); }
  return request;
};
const aborted = await run();
IDBObjectStore.prototype.put = original;
const afterAbort = await status();
const retried = await run();
engine.close();
db = await Database.open(database);
engine = new Engine(db,id);
const reopened = await engine.open();
const cells = await engine.read(reopened.snapshot.sheets[0].id,{startRow:0,endRow:65,startColumn:0,endColumn:1});
const scratch = await db.hasAny('scratch',scope);
engine.close();
postMessage({before,cancelled,cancelReachedScratch,sourceCancelled,afterCancel,rejected,afterRules,injected,aborted,afterAbort,retried,revision:reopened.revision,cells:cells.calculated,scratch});
}
`,
    }),
  );
  await page.goto('/validation-backup-harness');
}

test('restores mixed workbook metadata and formulas from a rule-free validation backup', async ({
  page,
}) => {
  await installHarness(page);
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-backup-test.worker.js?phase=metadata', {
          type: 'module',
        });
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
        worker.onerror = (error) => {
          worker.terminate();
          reject(new Error(error.message));
        };
      }),
  );
  expect(result).toMatchObject({
    dateSystem: '1904',
    extensions: { fixture: { provenance: 'mixed restore' } },
    names: [
      ['Data', false],
      ['Other', true],
    ],
    rows: { '0': { size: 42, hidden: true } },
    columns: { '1': { size: 140 } },
    freeze: { rows: 1, columns: 1 },
    merges: [{ startRow: 1, endRow: 2, startColumn: 3, endColumn: 5 }],
    first: {
      input: { type: 'number', value: 12 },
      note: 'source',
      link: { target: 'https://example.com/' },
    },
    style: { bold: true, numberFormat: '0.00' },
    formula: { input: { type: 'formula', expression: 'A1*2' } },
    value: 24,
  });
});

test('recovers a streamed rule-free validation backup in a real Worker', async ({ page }) => {
  await installHarness(page);
  const result = await page.evaluate(
    () =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker('/validation-backup-test.worker.js', { type: 'module' });
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
        worker.onerror = (error) => {
          worker.terminate();
          reject(new Error(error.message));
        };
      }),
  );
  expect(result.before).toMatchObject({ unchanged: true, scratch: false, revision: 0 });
  expect(result.afterRules).toEqual(result.before);
  expect(result.afterCancel).toEqual(result.before);
  expect(result.afterAbort).toEqual(result.before);
  expect(result).toMatchObject({
    rejected: 'UNSUPPORTED_FEATURE',
    cancelled: 'ABORTED',
    cancelReachedScratch: true,
    sourceCancelled: true,
    injected: true,
    aborted: 'ABORTED',
    retried: 'success',
    revision: 1,
    cells: { '0:0': 42, '64:0': 43 },
    scratch: false,
  });
});

for (const phase of ['scratch', 'heads'] as const) {
  test(`reopens a validation backup after Worker termination at ${phase} commit`, async ({
    page,
  }) => {
    await installHarness(page);
    const result = await page.evaluate(async (phase) => {
      const run = (params: Record<string, string>) =>
        new Promise<any>((resolve, reject) => {
          const worker = new Worker(
            `/validation-backup-test.worker.js?${new URLSearchParams(params)}`,
            { type: 'module' },
          );
          worker.onmessage = ({ data }) => {
            worker.terminate();
            resolve(data);
          };
          worker.onerror = (error) => {
            worker.terminate();
            reject(new Error(error.message));
          };
        });
      const interrupted = await run({ phase });
      return {
        reached: interrupted.phase,
        recovered: await run({
          phase: 'reopen',
          database: interrupted.database,
          id: interrupted.id,
        }),
      };
    }, phase);
    expect(result).toEqual({
      reached: phase,
      recovered: {
        revision: phase === 'heads' ? 1 : 0,
        values: phase === 'heads' ? [42, 43] : [null, null],
        scratch: false,
      },
    });
  });
}

test('measures a million-cell validation backup restore', async ({ page, browser }, info) => {
  const prefix = process.env.OPENSHEET_VALIDATION_RESTORE_REPORT;
  const format = process.env.OPENSHEET_RESTORE_FORMAT ?? 'validation-json';
  test.skip(!prefix, 'Opt-in performance workload');
  if (format !== 'validation-json' && format !== 'binary')
    throw new Error(`Unsupported OPENSHEET_RESTORE_FORMAT: ${format}`);
  test.setTimeout(120_000);
  await installHarness(page);
  const startedAt = new Date().toISOString();
  const result = await page.evaluate(
    (format) =>
      new Promise<any>((resolve, reject) => {
        const worker = new Worker(
          `/validation-backup-test.worker.js?phase=benchmark&format=${encodeURIComponent(format)}`,
          {
            type: 'module',
          },
        );
        worker.onmessage = ({ data }) => {
          worker.terminate();
          resolve(data);
        };
        worker.onerror = (error) => {
          worker.terminate();
          reject(new Error(error.message));
        };
      }),
    format,
  );
  writeFileSync(
    `${prefix}-${info.project.name}.json`,
    JSON.stringify(
      {
        browser: browser.version(),
        engine: info.project.name,
        cpu: cpus()[0]?.model,
        os: `${process.platform}/${process.arch}/${release()}`,
        source: sourceFingerprint(),
        startedAt,
        profile: 'Headless Vite source Worker; one sequential sample; no memory measurement',
        ...result,
      },
      null,
      2,
    ) + '\n',
    { flag: 'wx' },
  );
  expect(result).toMatchObject({
    format,
    cells: 1_000_000,
    pages: 628,
    revision: 1,
    verifiedBytes: true,
    verifiedCells: 1_000_000,
    verifiedLedger: true,
    values: [0, 6432, 999999],
    scratch: false,
  });
});
