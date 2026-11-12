import { expect, test } from '@playwright/test';

test('navigates qualified ranges and rejects invalid targets before switching sheets', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const ids = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const original = book.getSheets()[0].id;
    const target = await book.addSheet("O'Brien ! Q1", { rows: 10, columns: 10 });
    return { original, target: target.id };
  });
  const box = page.getByRole('textbox', { name: 'Cell address' });
  const initial = await page.evaluate(() => window.opensheet.selection.get());
  await box.fill("'O''Brien ! Q1'!A100");
  await box.press('Enter');
  await expect(page.locator('.os-toast')).toContainText('Range exceeds sheet dimensions');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(initial);
  await box.fill("'O''Brien ! Q1'!$B$2:$D$4");
  await box.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.selection.get()))
    .toEqual({ sheetId: ids.target, startRow: 1, endRow: 4, startColumn: 1, endColumn: 4 });
  await expect(box).toHaveValue('B2:D4');
  await box.fill('unfinished');
  await box.press('Escape');
  await expect(box).toHaveValue('B2:D4');
  await expect(page.getByRole('grid', { name: 'Spreadsheet', exact: true })).toBeFocused();
});

test('keeps an address draft during data refresh and ignores composing Enter', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const box = page.getByRole('textbox', { name: 'Cell address' });
  await box.fill('C10');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['refresh']]);
    const input = document.querySelector<HTMLInputElement>('.os-namebox')!;
    input.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Enter', isComposing: true, bubbles: true }),
    );
  });
  await expect(box).toHaveValue('C10');
  expect(await page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(0);
  await box.press('Enter');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(9);
});
