import { expect, test } from '@playwright/test';

test('depth exhaustion is not cached or hidden by IFERROR after reopening', async ({ page }) => {
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = await book.addSheet('Depth test', { rows: 130, columns: 1 });
    await sheet.range('A1').setValues([[1]]);
    await sheet
      .range('A2:A130')
      .setFormulas(
        Array.from({ length: 129 }, (_, i) => [i === 128 ? 'IFERROR(A129,99)' : `A${i + 1}+1`]),
      );
    await app.open(book.id);
    const reopened = app.getWorkbook().getSheetByName('Depth test')!;
    const limit = (await reopened.range('A130').getValues())[0][0];
    const withinLimit = (await reopened.range('A129').getValues())[0][0];
    const retried = (await reopened.range('A130').getValues())[0][0];
    return { limit, withinLimit, retried };
  });
  expect(result).toEqual({ limit: { error: '#LIMIT!' }, withinLimit: 129, retried: 129 });
});
