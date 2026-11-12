import { expect, test } from '@playwright/test';

test('finite large percentages retain their display after undo, redo and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const values = [1e308, -1e308, Number.MAX_VALUE, -Number.MAX_VALUE];
    const expected = values.map((value) => [
      new Intl.NumberFormat('en-US', {
        style: 'percent',
        useGrouping: true,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      }).format(value),
    ]);
    await sheet.range('A1:A4').setValues(values.map((value) => [value]));
    await sheet.range('A1:A4').setStyle({ numberFormat: '#,##0.00%' });
    const initial = await sheet.range('A1:A4').getDisplayValues();
    const undo = await book.undo(),
      redo = await book.redo();
    const redone = await sheet.range('A1:A4').getDisplayValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      expected,
      initial,
      undo,
      redo,
      redone,
      reopened: await sheet.range('A1:A4').getDisplayValues(),
      values: await sheet.range('A1:A4').getValues(),
    };
  });
  expect(result.undo).toBe(true);
  expect(result.redo).toBe(true);
  expect(result.initial).toEqual(result.expected);
  expect(result.redone).toEqual(result.expected);
  expect(result.reopened).toEqual(result.expected);
  expect(result.values).toEqual([[1e308], [-1e308], [Number.MAX_VALUE], [-Number.MAX_VALUE]]);
  expect(result.initial.flat().every((value) => !value.includes('∞'))).toBe(true);
});
