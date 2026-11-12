import { expect, test } from '@playwright/test';

test('inserts the maximum row count and preserves the original row through history', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 1, columns: 1 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[7]]);
    await sheet.setRowHeight(0, 30);
    const original = book.sheetData(sheet.id).rowOrder[0];
    await sheet.insertRows(0, 99999);
    const verify = () => {
      const data = book.sheetData(sheet.id);
      const order = data.rowOrder;
      return {
        count: order.length,
        idsCorrect: order.every((id, i) => id === (i === 99999 ? original : i + 1)),
        height: data.rows[original]?.size,
      };
    };
    const inserted = verify();
    const value = await sheet.range('A100000').getValues();
    await book.undo();
    const undone = {
      count: book.sheetData(sheet.id).rowOrder.length,
      value: await sheet.range('A1').getValues(),
    };
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      inserted,
      value,
      undone,
      reopened: verify(),
      reopenedValue: await sheet.range('A100000').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.inserted).toEqual({ count: 100000, idsCorrect: true, height: 30 });
  expect(result.value).toEqual([[7]]);
  expect(result.undone).toEqual({ count: 1, value: [[7]] });
  expect(result.reopened).toEqual(result.inserted);
  expect(result.reopenedValue).toEqual([[7]]);
});
