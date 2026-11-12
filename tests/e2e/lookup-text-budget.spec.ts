import { expect, test } from '@playwright/test';

test('literal and wildcard lookups bound long text across history and reopen', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const text = 'x'.repeat(32767);
    await sheet.range('A1:A30').setValues(Array.from({ length: 30 }, () => [text]));
    await sheet
      .range('C1:C3')
      .setFormulas([
        ['XLOOKUP("needle",A1:A31,B1:B31,"missing",0)'],
        ['MATCH("needle",A1:A31,0)'],
        ['XLOOKUP("n*",A1:A31,B1:B31,"missing",2)'],
      ]);
    const initial = await sheet.range('C1:C3').getValues();
    await sheet.range('A31').setValues([[text]]);
    const limited = await sheet.range('C1:C3').getValues();
    await book.undo();
    const undone = await sheet.range('C1:C3').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('C1:C3').getValues();
    await sheet.range('A31').setValues([['x']]);
    return {
      initial,
      limited,
      undone,
      reopened,
      recovered: await sheet.range('C1:C3').getValues(),
    };
  });
  const missing = [['missing'], [{ error: '#N/A' }], ['missing']];
  const limit = [[{ error: '#LIMIT!' }], [{ error: '#LIMIT!' }], [{ error: '#LIMIT!' }]];
  expect(result).toEqual({
    initial: missing,
    limited: limit,
    undone: missing,
    reopened: limit,
    recovered: missing,
  });
});
