import { test, expect } from '@playwright/test';

test('selective copy runs in the real Worker and persists formats, values and formulas', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:F4').clear({ all: true });
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('B1').setFormulas([['A1+$A$1']]);
    await sheet.range('A1:B1').setStyle({ bold: true, numberFormat: '0.00' });
    await sheet.range('A1:B1').copyTo('C1:D1', { mode: 'values' });
    const values = await sheet.range('C1:D1').getValues();
    const inputs = await sheet.range('C1:D1').getInputs();
    await sheet.range('A1:B1').copyTo('E1:F1', { mode: 'formulas' });
    const formulas = await sheet.range('E1:F1').getFormulas();
    await sheet.range('A1:B1').copyTo('C1:D1', { mode: 'formats' });
    const formatted = await sheet.range('C1:D1').getDisplayValues();
    await book.undo();
    const undone = await sheet.range('C1:D1').getDisplayValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      values,
      inputs,
      formulas,
      formatted,
      undone,
      reopened: await sheet.range('C1:D1').getDisplayValues(),
      formulaValues: await sheet.range('E1:F1').getValues(),
    };
  });
  expect(result).toEqual({
    values: [[7, 14]],
    inputs: [
      [
        { type: 'number', value: 7 },
        { type: 'number', value: 14 },
      ],
    ],
    formulas: [[null, '(E1+$A$1)']],
    formatted: [['7.00', '14.00']],
    undone: [['7', '14']],
    reopened: [['7.00', '14.00']],
    formulaValues: [[7, 14]],
  });
});
