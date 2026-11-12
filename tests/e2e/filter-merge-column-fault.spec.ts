import { expect, test } from '@playwright/test';

test('aborted multi-column deletion preserves filters, merges, formulas and history', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let abortColumnHead = false;
addEventListener('message', ({data}) => { if (data.abortColumnHead) abortColumnHead = true; });
const priorColumnPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = priorColumnPut.apply(this, args);
  if (abortColumnHead && this.name === 'heads') {
    abortColumnHead = false;
    this.transaction.abort();
    postMessage({columnHeadAborted: true});
  }
  return request;
};`,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(
      createWorkbookFile({
        sheets: [
          { name: 'Source', rows: 6, columns: 7 },
          { name: 'Other', rows: 2, columns: 4 },
        ],
      }),
    );
    let book = app.getWorkbook(),
      [source, other] = book.getSheets();
    const rows = Array.from({ length: 5 }, (_, i) => [
      i + 1,
      i % 2,
      i + 3,
      i + 4,
      i + 5,
      null,
      null,
    ]);
    await source.range('A1:G5').setValues(rows);
    await source.setFilter({
      mode: 'all',
      headerRows: 0,
      criteria: [
        { column: 0, operator: 'gt', value: 1 },
        { column: 1, operator: 'neq', value: 1 },
        { column: 3, operator: 'gt', value: 5 },
      ],
    });
    await other
      .range('A1:D1')
      .setFormulas([['Source!A1', 'Source!D1', 'SUM(Source!C1:D1)', 'SUM(Source!A1:D1)']]);
    await source.range('F6:G6').merge();
    const read = async () => {
      const meta = book.sheetData(source.id);
      return {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        ids: [...meta.columnOrder],
        nextId: meta.nextColumnId,
        filter: structuredClone(meta.filter),
        merges: structuredClone(meta.merges),
        hidden: book.filteredRows(source.id),
        cells: await source
          .range({ startRow: 0, endRow: 6, startColumn: 0, endColumn: meta.columnOrder.length })
          .getValues(),
        formulas: await other.range('A1:D1').getFormulas(),
        values: await other.range('A1:D1').getValues(),
      };
    };
    const before = await read(),
      worker = (book as any).worker as Worker;
    let injected = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.columnHeadAborted) injected = true;
    });
    worker.postMessage({ abortColumnHead: true });
    const code = await source.deleteColumns(0, 2).then(
      () => 'success',
      (error) => error.code ?? 'error',
    );
    const failed = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    const reopened = await read();
    const undoMerge = await book.undo(),
      withoutMerge = await read(),
      redoMerge = await book.redo(),
      restored = await read();
    await source.deleteColumns(0, 2);
    const deleted = await read();
    const undone = (await book.undo()) && (await read());
    const redone = (await book.redo()) && (await read());
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    return {
      injected,
      code,
      before,
      failed,
      reopened,
      undoMerge,
      withoutMerge,
      redoMerge,
      restored,
      deleted,
      undone,
      redone,
      final: await read(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.injected).toBe(true);
  expect(result.code).not.toBe('success');
  expect(result.failed).toEqual(result.before);
  expect(result.reopened).toEqual(result.before);
  expect(result.undoMerge).toBe(true);
  expect(result.withoutMerge.merges).toEqual([]);
  expect(result.withoutMerge).toMatchObject({
    ids: result.before.ids,
    filter: result.before.filter,
    hidden: result.before.hidden,
    cells: result.before.cells,
    formulas: result.before.formulas,
    values: result.before.values,
  });
  expect(result.redoMerge).toBe(true);
  expect(result.restored).toMatchObject({
    ids: result.before.ids,
    filter: result.before.filter,
    merges: result.before.merges,
    hidden: result.before.hidden,
    cells: result.before.cells,
    formulas: result.before.formulas,
    values: result.before.values,
  });
  expect(result.deleted.ids).toEqual([2, 3, 4, 5, 6]);
  expect(result.deleted.filter).toEqual({
    mode: 'all',
    headerRows: 0,
    criteria: [{ column: 1, operator: 'gt', value: 5 }],
  });
  expect(result.deleted.merges).toEqual([{ startRow: 5, startColumn: 3, endRow: 6, endColumn: 5 }]);
  expect(result.deleted.hidden).toEqual([0, 1, 5]);
  expect(result.deleted.cells).toEqual([
    ...Array.from({ length: 5 }, (_, i) => [i + 3, i + 4, i + 5, null, null]),
    [null, null, null, null, null],
  ]);
  expect(result.deleted.values).toEqual([[{ error: '#REF!' }, 4, 7, 7]]);
  expect(result.deleted.formulas[0][0]).toContain('#REF!');
  expect(result.deleted.formulas[0][1]).toContain('B1');
  expect(result.deleted.formulas[0][2]).toContain('A1:B1');
  expect(result.deleted.formulas[0][3]).toContain('A1:B1');
  expect(result.undone).toMatchObject({
    ids: result.before.ids,
    filter: result.before.filter,
    merges: result.before.merges,
    hidden: result.before.hidden,
    cells: result.before.cells,
    formulas: result.before.formulas,
    values: result.before.values,
  });
  expect(result.redone).toMatchObject({
    ids: result.deleted.ids,
    filter: result.deleted.filter,
    merges: result.deleted.merges,
    hidden: result.deleted.hidden,
    cells: result.deleted.cells,
    formulas: result.deleted.formulas,
    values: result.deleted.values,
  });
  expect(result.final).toEqual(result.redone);
});

test('reopens a committed multi-column deletion after an index failure without applying it twice', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let armColumnIndexFailure = false, failColumnIndex = false;
addEventListener('message', ({data}) => { if (data.armColumnIndexFailure) armColumnIndexFailure = true; });
const priorColumnCommitPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (armColumnIndexFailure && this.name === 'heads') {
    armColumnIndexFailure = false;
    const transaction = this.transaction, complete = transaction.oncomplete;
    transaction.oncomplete = function(event) {
      failColumnIndex = true;
      postMessage({columnFaultStage: 'committed'});
      if (complete) complete.call(this, event);
    };
  }
  return priorColumnCommitPut.apply(this, args);
};
const priorColumnMapSet = Map.prototype.set;
Map.prototype.set = function(...args) {
  if (failColumnIndex) {
    failColumnIndex = false;
    postMessage({columnFaultStage: 'index-failed'});
    throw new Error('Injected index failure after column commit');
  }
  return priorColumnMapSet.apply(this, args);
};`,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(
      createWorkbookFile({
        sheets: [
          { name: 'Source', rows: 6, columns: 7 },
          { name: 'Other', rows: 2, columns: 4 },
        ],
      }),
    );
    let book = app.getWorkbook(),
      [source, other] = book.getSheets();
    await source
      .range('A1:G5')
      .setValues(
        Array.from({ length: 5 }, (_, i) => [i + 1, i % 2, i + 3, i + 4, i + 5, null, null]),
      );
    await source.setFilter({
      mode: 'all',
      headerRows: 0,
      criteria: [
        { column: 0, operator: 'gt', value: 1 },
        { column: 1, operator: 'neq', value: 1 },
        { column: 3, operator: 'gt', value: 5 },
      ],
    });
    await other
      .range('A1:D1')
      .setFormulas([['Source!A1', 'Source!D1', 'SUM(Source!C1:D1)', 'SUM(Source!A1:D1)']]);
    await source.range('F6:G6').merge();
    const read = async () => {
      const meta = book.sheetData(source.id);
      return {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        ids: [...meta.columnOrder],
        nextId: meta.nextColumnId,
        filter: structuredClone(meta.filter),
        merges: structuredClone(meta.merges),
        hidden: book.filteredRows(source.id),
        cells: await source
          .range({ startRow: 0, endRow: 6, startColumn: 0, endColumn: meta.columnOrder.length })
          .getValues(),
        formulas: await other.range('A1:D1').getFormulas(),
        values: await other.range('A1:D1').getValues(),
      };
    };
    const before = await read(),
      stages: string[] = [],
      worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.columnFaultStage) stages.push(data.columnFaultStage);
    });
    worker.postMessage({ armColumnIndexFailure: true });
    const code = await source.deleteColumns(0, 2).then(
      () => 'success',
      (error) => error.code ?? 'error',
    );
    const state = book.persistenceState.state;
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    const recovered = await read();
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
    const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
    const requests = [
      tx.objectStore('heads').get(book.id),
      tx.objectStore('records').getAll(range),
      tx.objectStore('sizes').getAll(range),
      tx.objectStore('scratch').getAllKeys(range),
      tx.objectStore('records').getAllKeys(range),
      tx.objectStore('sizes').getAllKeys(range),
    ];
    const raw = await Promise.all(
      requests.map(
        (request) =>
          new Promise<any>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          }),
      ),
    );
    db.close();
    const rawBytes =
      raw[0].byteLength +
      [...raw[1], ...raw[2]].reduce(
        (total: number, value: Uint8Array) => total + value.byteLength,
        0,
      );
    const decoder = new TextDecoder();
    const ledger = raw[2].map((entry: Uint8Array) => JSON.parse(decoder.decode(entry.subarray(1))));
    const expectedLedger = raw[4].map((key: IDBValidKey, index: number) => [
      key,
      raw[1][index].byteLength,
    ]);
    const undo = await book.undo(),
      undone = await read(),
      redo = await book.redo(),
      redone = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    return {
      before,
      code,
      state,
      stages,
      recovered,
      rawBytes,
      scratch: raw[3].length,
      recordKeys: raw[4],
      sizeKeys: raw[5],
      ledger,
      expectedLedger,
      undo,
      undone,
      redo,
      redone,
      reopened: await read(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('WORKER_FAILED');
  expect(result.state).toBe('failed');
  expect(result.stages).toEqual(expect.arrayContaining(['committed', 'index-failed']));
  expect(result.recovered.revision).toBe(result.before.revision + 1);
  expect(result.recovered.ids).toEqual([2, 3, 4, 5, 6]);
  expect(result.recovered.nextId).toBe(result.before.nextId);
  expect(result.recovered.filter).toEqual({
    mode: 'all',
    headerRows: 0,
    criteria: [{ column: 1, operator: 'gt', value: 5 }],
  });
  expect(result.recovered.merges).toEqual([
    { startRow: 5, startColumn: 3, endRow: 6, endColumn: 5 },
  ]);
  expect(result.recovered.hidden).toEqual([0, 1, 5]);
  expect(result.recovered.cells).toEqual([
    ...Array.from({ length: 5 }, (_, i) => [i + 3, i + 4, i + 5, null, null]),
    [null, null, null, null, null],
  ]);
  expect(result.recovered.values).toEqual([[{ error: '#REF!' }, 4, 7, 7]]);
  expect(result.recovered.formulas[0][0]).toContain('#REF!');
  expect(result.recovered.formulas[0][1]).toContain('B1');
  expect(result.recovered.formulas[0][2]).toContain('A1:B1');
  expect(result.recovered.formulas[0][3]).toContain('A1:B1');
  expect(result.rawBytes).toBe(result.recovered.bytes);
  expect(result.scratch).toBe(0);
  expect(result.sizeKeys).toEqual(result.recordKeys);
  expect(result.ledger).toEqual(result.expectedLedger);
  expect(result.undo).toBe(true);
  expect(result.undone).toMatchObject({
    ids: result.before.ids,
    nextId: result.before.nextId,
    filter: result.before.filter,
    merges: result.before.merges,
    hidden: result.before.hidden,
    cells: result.before.cells,
    formulas: result.before.formulas,
    values: result.before.values,
  });
  expect(result.redo).toBe(true);
  expect(result.redone).toMatchObject({
    ids: result.recovered.ids,
    nextId: result.recovered.nextId,
    filter: result.recovered.filter,
    merges: result.recovered.merges,
    hidden: result.recovered.hidden,
    cells: result.recovered.cells,
    formulas: result.recovered.formulas,
    values: result.recovered.values,
  });
  expect(result.reopened).toEqual(result.redone);
});
