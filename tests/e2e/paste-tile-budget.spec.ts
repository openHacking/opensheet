import { expect, test } from '@playwright/test';

test('cumulative formula growth over the tile budget preserves persisted state', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchBudgetTile = false;
addEventListener('message', ({data}) => { if (data.watchBudgetTile) watchBudgetTile = true; });
const originalBudgetPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalBudgetPut.apply(this, args);
  if (watchBudgetTile && this.name === 'records') this.transaction.addEventListener('complete', () => {
    watchBudgetTile = false; postMessage({ budgetTilePageCommitted: true });
  }, { once: true });
  return request;
};`,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    const count = 1100;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Tile', rows: count, columns: 1 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[17]]);
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const request = <T>(r: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    const raw = async () => {
      const db = await request(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const prefix = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          request(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            request(tx.objectStore(name).getAllKeys(prefix)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const baseline = await raw();
    const worker = (book as any).worker as Worker;
    let staged = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.budgetTilePageCommitted) staged = true;
    });
    worker.postMessage({ watchBudgetTile: true });
    const item = {
      input: { type: 'formula', expression: 'A1+A1' },
      source: { row: 0, column: 0 },
      note: '',
    };
    const budget = 32 * 1024 * 1024;
    item.note = 'x'.repeat(Math.floor((budget - 1) / count) - JSON.stringify(item).length - 34);
    const base = (JSON.stringify(item).length + 32) * count;
    const failure = await book
      .execute({
        type: 'core.cells.pasteTile',
        payload: {
          sheetId: sheet.id,
          range: { startRow: 0, endRow: count, startColumn: 0, endColumn: 1 },
          rows: 1,
          columns: 1,
          cells: [item],
        },
      })
      .then(
        () => ({ code: 'success', message: '' }),
        (error) => ({ code: error.code, message: error.message }),
      );
    const after = await raw();
    const current = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      values: await sheet.range('A1:A2').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1:A2').getValues();
    await book.undo();
    const undone = await sheet.range('A1').getValues();
    await book.redo();
    return {
      staged,
      base,
      budget,
      failure,
      baseline,
      after,
      revision,
      bytes,
      current,
      reopened,
      undone,
      redone: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.base).toBeLessThanOrEqual(result.budget);
  expect(result.staged).toBe(true);
  expect(result.failure).toEqual({
    code: 'LIMIT_EXCEEDED',
    message: 'Expanded clipboard content is too large',
  });
  expect(result.after).toEqual(result.baseline);
  expect(result.current).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    values: [[17], [null]],
  });
  expect(result.reopened).toEqual([[17], [null]]);
  expect(result.undone).toEqual([[null]]);
  expect(result.redone).toEqual([[17]]);
});
