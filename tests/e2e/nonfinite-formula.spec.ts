import { expect, test } from '@playwright/test';

test('overflow formula survives copy, row insertion, history and reopen', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = await book.addSheet('Overflow', { rows: 5, columns: 2 });
    await sheet.range('A1:A2').setValues([[5], [9]]);
    await sheet.range('B1').setFormulas([['IFERROR(1e309,A1)']]);
    const initial = await sheet.range('B1').getValues();
    await sheet.range('B1').copyTo('B2');
    const copied = await sheet.range('B2').getValues();
    const copiedFormula = await sheet.range('B2').getFormulas();
    await sheet.insertRows(0);
    const inserted = await sheet.range('B3').getValues();
    await book.undo();
    const undone = await sheet.range('B2').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheetByName('Overflow')!;
    return {
      initial,
      copied,
      copiedFormula,
      inserted,
      undone,
      reopened: await sheet.range('B3').getValues(),
      formula: await sheet.range('B3').getFormulas(),
    };
  });
  expect(result.initial).toEqual([[5]]);
  for (const phase of [result.copied, result.inserted, result.undone, result.reopened])
    expect(phase).toEqual([[9]]);
  expect(result.copiedFormula[0][0]).toContain('#NUM!');
  expect(result.formula[0][0]).toContain('#NUM!');
});
