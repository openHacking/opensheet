import { expect, test } from '@playwright/test';

test('date-time values retain their time in the rendered and copied display after reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[45293.5]]);
    await sheet.range('A1').setStyle({ numberFormat: 'yyyy-mm-dd hh:mm' });
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    const display = await sheet.range('A1').getDisplayValues();
    const copied = await app.getGrid()!.copyClipboard();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      display,
      copied: copied.text,
      reopened: await sheet.range('A1').getDisplayValues(),
      value: await sheet.range('A1').getValues(),
    };
  });
  expect(result).toEqual({
    display: [['2024-01-02 12:00']],
    copied: '2024-01-02 12:00',
    reopened: [['2024-01-02 12:00']],
    value: [[45293.5]],
  });
});
