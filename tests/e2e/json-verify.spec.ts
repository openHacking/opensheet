import { expect, test } from '@playwright/test';

test('verifies native JSON without replacing the current workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (moduleUrl) => {
    const { verifyJSON } = await import(/* @vite-ignore */ moduleUrl);
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[987]]);
    const revision = book.revision;
    const text = JSON.stringify(await book.toJSON());
    const verified = await verifyJSON(new Blob([text]));
    let rejected = false;
    try {
      await verifyJSON(text.slice(0, -1));
    } catch {
      rejected = true;
    }
    return {
      verified,
      rejected,
      before: revision,
      after: book.revision,
      value: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.verified.workbookId).toBeTruthy();
  expect(result.verified.pages).toBeGreaterThan(0);
  expect(result.verified.cells).toBeGreaterThan(0);
  expect(result.rejected).toBe(true);
  expect(result.after).toBe(result.before);
  expect(result.value).toEqual([[987]]);
});
