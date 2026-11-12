import { expect, test } from '@playwright/test';

test('direct paste rejects malformed quoted text without a write or selection change', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['original']]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    (window as any).malformedPasteErrors = [];
    (app.getGrid() as any).options.onError = (error: Error) =>
      (window as any).malformedPasteErrors.push(error.message);
    return { revision: book.revision, selection: app.selection.get() };
  });
  await page.getByRole('grid', { name: 'Spreadsheet' }).evaluate((element) => {
    const clipboard = new DataTransfer();
    clipboard.setData('text/plain', '"new"tail');
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboard });
    element.dispatchEvent(event);
  });
  await expect
    .poll(() => page.evaluate(() => (window as any).malformedPasteErrors))
    .toEqual(['Unexpected text after closing quote']);
  const after = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    return {
      revision: book.revision,
      selection: app.selection.get(),
      value: await book.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(after).toEqual({ ...before, value: [['original']] });
});
