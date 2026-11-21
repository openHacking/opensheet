import { expect, test } from '@playwright/test';
import { createSheet } from '../../packages/core/src/model.js';

test('adding a populated sheet persists its cells through history and reopen', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Seeded', 70, 2);
  initial.cells['0:0'] = {
    rowId: 0,
    columnId: 0,
    input: { type: 'number', value: 4 },
    note: 'seed',
  };
  initial.cells['64:0'] = { rowId: 64, columnId: 0, input: { type: 'number', value: 7 } };
  initial.cells['64:1'] = {
    rowId: 64,
    columnId: 1,
    input: { type: 'formula', expression: 'A1+A65' },
  };
  const result = await page.evaluate(async (inputJSON) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const source = book.getSheets()[0];
    await source.range('A30').setStyle({ bold: true });
    const styleId = (await book.getCell(source.id, 29, 0))!.styleId;
    const sheet = JSON.parse(inputJSON);
    sheet.cells['0:0'].styleId = styleId;
    const envelope = {
      protocolVersion: 1 as const,
      commandId: 'seeded-sheet-retry',
      workbookId: book.id,
      baseRevision: book.revision,
      type: 'core.sheet.add',
      payload: { sheet },
    };
    const commit = await book.execute(envelope);
    const addedRevision = book.revision;
    const duplicate = await book.execute(envelope);
    const retryUnchanged =
      book.revision === addedRevision && JSON.stringify(duplicate) === JSON.stringify(commit);
    const commandCells = (commit!.commands[0].payload.sheet as typeof sheet).cells;
    const inspect = async () => {
      const added = book.getSheetById(sheet.id)!;
      return {
        values: [await added.range('A1').getValues(), await added.range('A65:B65').getValues()],
        formulas: await added.range('B65').getFormulas(),
        note: await added.range('A1').getNotes(),
        style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
      };
    };
    const added = await inspect();
    await book.undo();
    const removed = book.getSheetById(sheet.id) === undefined;
    await book.redo();
    const redone = await inspect();
    await app.open(book.id);
    book = app.getWorkbook();
    return { added, removed, redone, reopened: await inspect(), commandCells, retryUnchanged };
  }, JSON.stringify(initial));
  const expected = {
    values: [[[4]], [[7, 11]]],
    formulas: [['A1+A65']],
    note: [['seed']],
    style: { bold: true },
  };
  expect(result.added).toEqual(expected);
  expect(result.removed).toBe(true);
  expect(result.redone).toEqual(expected);
  expect(result.reopened).toEqual(expected);
  expect(result.retryUnchanged).toBe(true);
  expect(Object.keys(result.commandCells).sort()).toEqual(['0:0', '64:0', '64:1']);
  expect(result.commandCells['64:1'].input).toEqual({ type: 'formula', expression: 'A1+A65' });
});

test('transaction commands can edit a newly added populated sheet in order', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Seeded', 70, 2);
  initial.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } };
  initial.cells['64:0'] = { rowId: 64, columnId: 0, input: { type: 'number', value: 7 } };
  const result = await page.evaluate(async (inputJSON) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const initial = JSON.parse(inputJSON);
    await book.transaction({}, async (tx) => {
      await tx.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
      const sheet = tx.getSheetById(initial.id)!;
      await sheet.range('B1').setFormulas([['A1+A65']]);
      await sheet.range('A65').setValues([[8]]);
    });
    const read = async () => book.getSheetById(initial.id)!.range('A1:B1').getValues();
    const added = await read();
    await book.undo();
    const removed = book.getSheetById(initial.id) === undefined;
    await book.redo();
    const redone = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    return { added, removed, redone, reopened: await read() };
  }, JSON.stringify(initial));
  expect(result).toEqual({
    added: [[4, 12]],
    removed: true,
    redone: [[4, 12]],
    reopened: [[4, 12]],
  });
});

test('a missing prototype-named style rejects the complete transaction', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Seeded', 2, 2);
  initial.cells['0:0'] = {
    rowId: 0,
    columnId: 0,
    input: { type: 'number', value: 4 },
    styleId: 'toString',
  };
  const result = await page.evaluate(async (inputJSON) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A30').setValues([[20]]);
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const initial = JSON.parse(inputJSON);
    const error = await book
      .transaction({}, async (tx) => {
        await tx
          .getSheetById(sheet.id)!
          .range('A30')
          .setValues([[99]]);
        await tx.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
      })
      .then(
        () => 'unexpected success',
        (failure: { code: string }) => failure.code,
      );
    const failed = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      value: await sheet.range('A30').getValues(),
      missing: !book.getSheetById(initial.id),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = { revision: book.revision, value: await sheet.range('A30').getValues() };
    await book.undo();
    const undone = await sheet.range('A30').getValues();
    await book.redo();
    return {
      error,
      revision,
      bytes,
      failed,
      reopened,
      undone,
      redone: await sheet.range('A30').getValues(),
    };
  }, JSON.stringify(initial));
  expect(result.error).toBe('INVALID_ARGUMENT');
  expect(result.failed).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    value: [[20]],
    missing: true,
  });
  expect(result.reopened).toEqual({ revision: result.revision, value: [[20]] });
  expect(result.undone).toEqual([[null]]);
  expect(result.redone).toEqual([[20]]);
});

test('oversized initial cell content is rejected before publication and can be retried', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Seeded', 300, 1);
  const result = await page.evaluate(async (inputJSON) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const initial = JSON.parse(inputJSON);
    const text = 'x'.repeat(30_000);
    for (let row = 0; row < 280; row++)
      initial.cells[`${row}:0`] = {
        rowId: row,
        columnId: 0,
        input: { type: 'string', value: text },
      };
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const error = await book.execute({ type: 'core.sheet.add', payload: { sheet: initial } }).then(
      () => ({ code: 'unexpected success', message: '' }),
      (failure: { code: string; message: string }) => ({
        code: failure.code,
        message: failure.message,
      }),
    );
    const failed = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      missing: !book.getSheetById(initial.id),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    const reopened = { revision: book.revision, missing: !book.getSheetById(initial.id) };
    initial.cells = { '0:0': { rowId: 0, columnId: 0, input: { type: 'string', value: 'small' } } };
    await book.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
    return {
      error,
      revision,
      bytes,
      failed,
      reopened,
      retry: await book.getSheetById(initial.id)!.range('A1').getValues(),
    };
  }, JSON.stringify(initial));
  expect(result.error).toMatchObject({
    code: 'LIMIT_EXCEEDED',
    message: expect.stringContaining('Initial sheet cells'),
  });
  expect(result.failed).toEqual({ revision: result.revision, bytes: result.bytes, missing: true });
  expect(result.reopened).toEqual({ revision: result.revision, missing: true });
  expect(result.retry).toEqual([['small']]);
});

test('initial sheet cell count is bounded before publication', async ({ page }) => {
  test.setTimeout(90_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Seeded', 656, 100);
  const result = await page.evaluate(async (inputJSON) => {
    const book = window.opensheet.getWorkbook();
    const initial = JSON.parse(inputJSON);
    for (let position = 0; position < 65_537; position++) {
      const row = Math.floor(position / 100),
        column = position % 100;
      initial.cells[`${row}:${column}`] = {
        rowId: row,
        columnId: column,
        input: { type: 'number', value: 1 },
      };
    }
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const error = await book.execute({ type: 'core.sheet.add', payload: { sheet: initial } }).then(
      () => ({ code: 'unexpected success', message: '' }),
      (failure: { code: string; message: string }) => ({
        code: failure.code,
        message: failure.message,
      }),
    );
    return {
      error,
      revision,
      after: book.revision,
      bytes,
      afterBytes: (await book.storageStats()).bytes,
      missing: !book.getSheetById(initial.id),
    };
  }, JSON.stringify(initial));
  expect(result.error).toMatchObject({
    code: 'LIMIT_EXCEEDED',
    message: expect.stringContaining('Initial sheet cells'),
  });
  expect(result.after).toBe(result.revision);
  expect(result.afterBytes).toBe(result.bytes);
  expect(result.missing).toBe(true);
});

test('the initial sheet count boundary persists every cell through undo and redo', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Seeded', 656, 100);
  const result = await page.evaluate(async (inputJSON) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const initial = JSON.parse(inputJSON);
    for (let position = 0; position < 65_536; position++) {
      const row = Math.floor(position / 100),
        column = position % 100;
      initial.cells[`${row}:${column}`] = {
        rowId: row,
        columnId: column,
        input: { type: 'number', value: 1 },
      };
    }
    await book.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
    const verify = async () => {
      const sheet = book.getSheetById(initial.id)!;
      let checked = 0;
      for (let row = 0; row < 656; row += 100) {
        const values = await sheet
          .range({
            startRow: row,
            endRow: Math.min(row + 100, 656),
            startColumn: 0,
            endColumn: 100,
          })
          .getValues();
        for (const [offset, cells] of values.entries())
          for (const [column, value] of cells.entries()) {
            const position = (row + offset) * 100 + column;
            if (value !== (position < 65_536 ? 1 : null))
              throw new Error(`Mismatch at ${position}`);
            checked++;
          }
      }
      return checked;
    };
    const added = await verify();
    await book.undo();
    const undone = !book.getSheetById(initial.id);
    await book.redo();
    const redone = await verify();
    await app.open(book.id);
    book = app.getWorkbook();
    return { added, undone, redone, reopened: await verify() };
  }, JSON.stringify(initial));
  expect(result).toEqual({ added: 65_600, undone: true, redone: 65_600, reopened: 65_600 });
});

test('replacing a removed sheet ID in one transaction preserves old and new pages in history', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = createSheet('Replacement', 70, 2);
  initial.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } };
  initial.cells['64:0'] = { rowId: 64, columnId: 0, input: { type: 'number', value: 7 } };
  const result = await page.evaluate(async (inputJSON) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const old = await book.addSheet('Old', { rows: 70, columns: 2 });
    await old.range('A1').setValues([[99]]);
    await old.range('A65').setValues([[100]]);
    const initial = JSON.parse(inputJSON);
    initial.id = old.id;
    await book.transaction({}, async (tx) => {
      await tx.removeSheet(old.id);
      await tx.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
    });
    const read = async () => {
      const sheet = book.getSheetById(old.id)!;
      return {
        name: sheet.name,
        values: [await sheet.range('A1').getValues(), await sheet.range('A65').getValues()],
      };
    };
    const replaced = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    const redone = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    return { replaced, undone, redone, reopened: await read() };
  }, JSON.stringify(initial));
  const replaced = { name: 'Replacement', values: [[[4]], [[7]]] };
  expect(result).toEqual({
    replaced,
    undone: { name: 'Old', values: [[[99]], [[100]]] },
    redone: replaced,
    reopened: replaced,
  });
});

for (const fault of ['page', 'head'] as const) {
  test(`initial-sheet ${fault} failure rolls back staged records and can be retried`, async ({
    page,
  }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let watchInitialSheet = false, initialSheetWrites = 0;
const initialSheetFault = ${JSON.stringify(fault)};
addEventListener('message', ({data}) => { if (data.watchInitialSheet) watchInitialSheet = true; });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (initialSheetFault === 'page' && watchInitialSheet && this.name === 'records' && initialSheetWrites === 1) {
    watchInitialSheet = false;
    postMessage({initialSheetWritesBeforeFailure: initialSheetWrites});
    throw new DOMException('Injected page failure', 'QuotaExceededError');
  }
  const request = originalPut.apply(this, args);
  if (initialSheetFault === 'head' && watchInitialSheet && this.name === 'heads') {
    const transaction = this.transaction;
    request.addEventListener('success', () => {
      watchInitialSheet = false;
      transaction.abort();
      postMessage({initialSheetWritesBeforeFailure: initialSheetWrites});
    }, {once: true});
  }
  if (watchInitialSheet && this.name === 'records')
    this.transaction.addEventListener('complete', () => { initialSheetWrites++; }, {once: true});
  return request;
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const initial = createSheet('Seeded', 70, 2);
    initial.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } };
    initial.cells['64:0'] = { rowId: 64, columnId: 0, input: { type: 'number', value: 7 } };
    const result = await page.evaluate(async (inputJSON) => {
      const app = window.opensheet;
      let book = app.getWorkbook();
      const initial = JSON.parse(inputJSON);
      const get = <T>(request: IDBRequest<T>) =>
        new Promise<T>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
      const raw = async () => {
        const db = await get(indexedDB.open('opensheet'));
        try {
          const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
          const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
          const [head, keys] = await Promise.all([
            get(tx.objectStore('heads').get(book.id)),
            Promise.all(
              ['records', 'sizes', 'scratch'].map((name) =>
                get(tx.objectStore(name).getAllKeys(range)),
              ),
            ),
          ]);
          return { head: Array.from(head as Uint8Array), keys };
        } finally {
          db.close();
        }
      };
      const before = {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        raw: await raw(),
      };
      const worker = (book as any).worker as Worker;
      let writes = 0;
      worker.addEventListener('message', ({ data }) => {
        if (data.initialSheetWritesBeforeFailure) writes = data.initialSheetWritesBeforeFailure;
      });
      worker.postMessage({ watchInitialSheet: true });
      const error = await book
        .execute({ type: 'core.sheet.add', payload: { sheet: initial } })
        .then(
          () => 'unexpected success',
          (failure: { code: string }) => failure.code,
        );
      const failed = {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        raw: await raw(),
        missing: !book.getSheetById(initial.id),
      };
      await app.open(book.id);
      book = app.getWorkbook();
      const reopened = { revision: book.revision, missing: !book.getSheetById(initial.id) };
      await book.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
      const read = async () => {
        const sheet = book.getSheetById(initial.id)!;
        return [await sheet.range('A1').getValues(), await sheet.range('A65').getValues()];
      };
      const retry = await read();
      await book.undo();
      const undone = !book.getSheetById(initial.id);
      await book.redo();
      const redone = await read();
      return { writes, error, before, failed, reopened, retry, undone, redone };
    }, JSON.stringify(initial));
    if (fault === 'page') {
      expect(result.writes).toBe(1);
      expect(result.error).toBe('QuotaExceededError');
    } else {
      expect(result.writes).toBeGreaterThanOrEqual(2);
      expect(result.error).not.toBe('unexpected success');
    }
    expect(result.failed).toEqual({ ...result.before, missing: true });
    expect(result.reopened).toEqual({ revision: result.before.revision, missing: true });
    expect(result.retry).toEqual([[[4]], [[7]]]);
    expect(result.undone).toBe(true);
    expect(result.redone).toEqual(result.retry);
  });
}
