import { expect, test } from '@playwright/test';

test('Shift+Home stays within the cell editor while editing', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('E5').setValues([['stored']]);
    app.selection.set({ sheetId: sheet.id, startRow: 4, endRow: 5, startColumn: 4, endColumn: 5 });
  });
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await grid.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await editor.fill('abcdef');
  await editor.evaluate((element) => (element as HTMLTextAreaElement).setSelectionRange(4, 4));
  await editor.press('Shift+Home');
  expect(
    await editor.evaluate((element) => [
      (element as HTMLTextAreaElement).selectionStart,
      (element as HTMLTextAreaElement).selectionEnd,
    ]),
  ).toEqual([0, 4]);
  expect(
    await page.evaluate(async () => ({
      selection: window.opensheet.selection.get(),
      value: await window.opensheet.getWorkbook().getSheets()[0].range('E5').getValues(),
    })),
  ).toMatchObject({
    selection: { startRow: 4, endRow: 5, startColumn: 4, endColumn: 5 },
    value: [['stored']],
  });
});
