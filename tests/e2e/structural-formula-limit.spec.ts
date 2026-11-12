import { expect, test } from '@playwright/test';

test('sheet rename rejects an oversized rewritten formula and preserves history', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(
      createWorkbookFile({
        sheets: [
          { name: 'S', rows: 2, columns: 2 },
          { name: 'Other', rows: 2, columns: 2 },
        ],
      }),
    );
    let book = app.getWorkbook();
    const [source, other] = book.getSheets();
    const prefix = 'S!A1&"';
    const expression = prefix + 'x'.repeat(32768 - prefix.length - 1) + '"';
    await other.range('A1:B1').setFormulas([['S!A1', expression]]);
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const code = await book.renameSheet(source.id, 'LongerName').then(
      () => 'success',
      (error) => error.code,
    );
    const current = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      name: book.getSheets()[0].name,
      formulas: await other.range('A1:B1').getFormulas(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    const reopened = {
      name: book.getSheets()[0].name,
      formulas: await book.getSheets()[1].range('A1:B1').getFormulas(),
    };
    await book.undo();
    await book.redo();
    return {
      code,
      revision,
      bytes,
      current,
      reopened,
      redone: await book.getSheets()[1].range('A1:B1').getFormulas(),
      expression,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.current).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    name: 'S',
    formulas: [['S!A1', result.expression]],
  });
  expect(result.reopened).toEqual({ name: 'S', formulas: result.current.formulas });
  expect(result.redone).toEqual(result.current.formulas);
});

test('late axis rewrite failure cleans up already staged formula pages', async ({ page }) => {
  test.setTimeout(60_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchRewrite = false;
addEventListener('message', ({data}) => { if (data.watchRewrite) watchRewrite = true; });
const originalRewritePut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalRewritePut.apply(this, args);
  if (watchRewrite && this.name === 'records') this.transaction.addEventListener('complete', () => {
    watchRewrite = false; postMessage({ rewritePageCommitted: true });
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
    await app.load(createWorkbookFile({ sheets: [{ name: 'S', rows: 130, columns: 1 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const expression = `A9&"${'x'.repeat(32763)}"`;
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [
          { row: 0, column: 0, input: { type: 'formula', expression: 'A9' } },
          { row: 64, column: 0, input: { type: 'formula', expression } },
        ],
      },
    });
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
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const worker = (book as any).worker as Worker;
    let staged = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.rewritePageCommitted) staged = true;
    });
    worker.postMessage({ watchRewrite: true });
    const code = await sheet.insertRows(0).then(
      () => 'success',
      (error) => error.code,
    );
    const after = await raw();
    const current = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      rows: book.sheetData(sheet.id).rowOrder.length,
      first: await sheet.range('A1').getFormulas(),
      last: await sheet.range('A65').getFormulas(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = {
      rows: book.sheetData(sheet.id).rowOrder.length,
      first: await sheet.range('A1').getFormulas(),
      last: await sheet.range('A65').getFormulas(),
    };
    await book.undo();
    await book.redo();
    const redone = await sheet.range('A65').getFormulas();
    await sheet.range('A65').clear();
    await sheet.insertRows(0);
    const retry = {
      rows: book.sheetData(sheet.id).rowOrder.length,
      formula: await sheet.range('A2').getFormulas(),
    };
    await book.undo();
    const undoRows = book.sheetData(sheet.id).rowOrder.length;
    await book.redo();
    const redoRows = book.sheetData(sheet.id).rowOrder.length;
    return {
      staged,
      code,
      baseline,
      after,
      revision,
      bytes,
      current,
      reopened,
      redone,
      retry,
      undoRows,
      redoRows,
      expression,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.staged).toBe(true);
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.after).toEqual(result.baseline);
  expect(result.current).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    rows: 130,
    first: [['A9']],
    last: [[result.expression]],
  });
  expect(result.reopened).toEqual({ rows: 130, first: [['A9']], last: [[result.expression]] });
  expect(result.redone).toEqual([[result.expression]]);
  expect(result.retry).toEqual({ rows: 131, formula: [["'S'!A10"]] });
  expect(result.undoRows).toBe(130);
  expect(result.redoRows).toBe(131);
});
