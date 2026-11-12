import { expect, test } from '@playwright/test';

test('200% grid zoom survives workbook reopening and replacement', async ({ page }) => {
  await page.setViewportSize({ width: 900, height: 800 });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.getByLabel('Zoom').selectOption('2');
  expect(await page.evaluate(() => window.opensheet.getGrid()!.getZoom())).toBe(2);
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await expect(editor).toBeVisible();
  await editor.fill('zoom draft');
  await expect
    .poll(() =>
      editor.evaluate((element) => {
        const rect = element.getBoundingClientRect(),
          parent = element.parentElement!.getBoundingClientRect();
        return (
          rect.left >= parent.left &&
          rect.right <= parent.right + 1 &&
          rect.bottom <= parent.bottom + 1
        );
      }),
    )
    .toBe(true);
  await editor.press('Escape');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const id = app.getWorkbook().id;
    const errors: string[] = [];
    for (const value of [NaN, Infinity, -Infinity]) {
      try {
        app.getGrid()!.setZoom(value);
      } catch (error) {
        errors.push((error as { code: string }).code);
      }
    }
    const afterInvalid = app.getGrid()!.getZoom();
    await app.open(id);
    return { errors, afterInvalid, reopened: app.getGrid()!.getZoom() };
  });
  expect(result).toEqual({
    errors: ['INVALID_ARGUMENT', 'INVALID_ARGUMENT', 'INVALID_ARGUMENT'],
    afterInvalid: 2,
    reopened: 2,
  });
  await page.getByLabel('Choose demo').selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await expect(page.getByLabel('Zoom')).toHaveValue('2');
  expect(await page.evaluate(() => window.opensheet.getGrid()!.getZoom())).toBe(2);
});
