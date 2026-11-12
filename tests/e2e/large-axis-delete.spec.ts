import { expect, test } from '@playwright/test';

test('deletes 99,999 populated rows atomically and restores every value through history', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 100000, columns: 1 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:A100000').fillSeries({ start: 1, step: 1 });
    await sheet.setRowHeight(99999, 30);
    await sheet.setRowHidden(50000, true);
    const id = sheet.id;
    const lastId = book.sheetData(id).rowOrder[99999];
    const initialRevision = book.revision;
    await sheet.deleteRows(0, 99999);
    const deleted = {
      count: book.sheetData(id).rowOrder.length,
      lastId: book.sheetData(id).rowOrder[0],
      height: book.sheetData(id).rows[lastId]?.size,
      value: await sheet.range('A1').getValues(),
    };
    await book.undo();
    let allRestored = true;
    for (let startRow = 0; startRow < 100000; startRow += 10000) {
      const values = await sheet
        .range({ startRow, endRow: startRow + 10000, startColumn: 0, endColumn: 1 })
        .getValues();
      if (values.some((row, i) => row[0] !== startRow + i + 1)) allRestored = false;
    }
    const undone = {
      count: book.sheetData(id).rowOrder.length,
      hidden: book.sheetData(id).rows[50000]?.hidden,
      allRestored,
    };
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const data = book.sheetData(id);
    return {
      initialRevision,
      finalRevision: book.revision,
      deleted,
      undone,
      reopened: {
        count: data.rowOrder.length,
        lastId: data.rowOrder[0],
        height: data.rows[lastId]?.size,
        value: await sheet.range('A1').getValues(),
      },
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.deleted).toEqual({ count: 1, lastId: 99999, height: 30, value: [[100000]] });
  expect(result.undone).toEqual({ count: 100000, hidden: true, allRestored: true });
  expect(result.reopened).toEqual(result.deleted);
  expect(result.finalRevision).toBeGreaterThan(result.initialRevision);
});

test('aborted head publication preserves pages removed by an axis deletion', async ({ page }) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let abortDeleteHead = false;
addEventListener('message', ({data}) => { if (data.abortDeleteHead) abortDeleteHead = true; });
const originalDeletePut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalDeletePut.apply(this, args);
  if (abortDeleteHead && this.name === 'heads') {
    abortDeleteHead = false;
    this.transaction.abort();
    postMessage({deleteHeadAborted: true});
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
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 130, columns: 1 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:A130').fillSeries({ start: 1, step: 1 });
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const worker = (book as any).worker as Worker;
    let injected = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.deleteHeadAborted) injected = true;
    });
    worker.postMessage({ abortDeleteHead: true });
    const code = await sheet.deleteRows(0, 70).then(
      () => 'success',
      (error) => error.code ?? 'error',
    );
    const current = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      values: await sheet.range('A1:A130').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1:A130').getValues();
    await sheet.deleteRows(0, 70);
    const retried = await sheet.range('A1:A60').getValues();
    await book.undo();
    const undone = await sheet.range('A1:A130').getValues();
    await book.redo();
    return {
      injected,
      code,
      revision,
      bytes,
      current,
      reopened,
      retried,
      undone,
      redone: await sheet.range('A1:A60').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const initial = Array.from({ length: 130 }, (_, i) => [i + 1]);
  const remaining = Array.from({ length: 60 }, (_, i) => [i + 71]);
  expect(result.injected).toBe(true);
  expect(result.code).not.toBe('success');
  expect(result.current).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    values: initial,
  });
  expect(result.reopened).toEqual(initial);
  expect(result.retried).toEqual(remaining);
  expect(result.undone).toEqual(initial);
  expect(result.redone).toEqual(remaining);
});

test('deleting a complete page and part of another preserves survivor references', async ({
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
          { name: 'Source', rows: 130, columns: 1 },
          { name: 'Other', rows: 2, columns: 2 },
        ],
      }),
    );
    let book = app.getWorkbook();
    let [source, other] = book.getSheets();
    await source.range('A1:A130').fillSeries({ start: 1, step: 1 });
    await other.range('A1:B1').setFormulas([['Source!A1', 'Source!A71']]);
    await source.deleteRows(0, 70);
    const deleted = {
      values: await source.range('A1:A60').getValues(),
      formulas: await other.range('A1:B1').getFormulas(),
      results: await other.range('A1:B1').getValues(),
    };
    await book.undo();
    const undone = {
      values: await source.range('A1:A130').getValues(),
      formulas: await other.range('A1:B1').getFormulas(),
      results: await other.range('A1:B1').getValues(),
    };
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    const reopened = {
      values: await source.range('A1:A60').getValues(),
      formulas: await other.range('A1:B1').getFormulas(),
      results: await other.range('A1:B1').getValues(),
    };
    return { deleted, undone, reopened };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.deleted.values).toEqual(Array.from({ length: 60 }, (_, i) => [i + 71]));
  expect(result.deleted.results).toEqual([[{ error: '#REF!' }, 71]]);
  expect(result.undone.values).toEqual(Array.from({ length: 130 }, (_, i) => [i + 1]));
  expect(result.undone.results).toEqual([[1, 71]]);
  expect(result.reopened).toEqual(result.deleted);
});
