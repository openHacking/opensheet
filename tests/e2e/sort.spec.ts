import { workbookTool } from './workbook-tools.js';
import { test, expect } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    await app.load(JSON.parse(fileJSON));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:C5').setValues([
      ['Group', 'Score', 'ID'],
      ['b', 1, 'first'],
      ['a', 1, 'second'],
      ['a', 3, 'third'],
      ['a', 3, 'fourth'],
    ]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 });
  }, JSON.stringify(createWorkbookFile()));
});

test('custom sort preserves headers, stable ties, selection and history', async ({ page }) => {
  await workbookTool(page, 'sort-custom');
  await expect(page.getByRole('dialog')).toBeVisible();
  await page.getByLabel('First selected row contains headers').check();
  await page.getByRole('button', { name: 'Add sort level' }).click();
  await page.getByLabel('Level 2 order').selectOption('desc');
  await page.getByRole('button', { name: 'Apply sort' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const sorted = await sheet.range('A1:C5').getValues();
    const selection = app.selection.get();
    await book.undo();
    const original = await sheet.range('C2:C5').getValues();
    await book.redo();
    await app.open(book.id);
    return {
      sorted,
      selection,
      original,
      restored: await app.getWorkbook().getSheets()[0].range('A1:C5').getValues(),
    };
  });
  expect(result.sorted).toEqual([
    ['Group', 'Score', 'ID'],
    ['a', 3, 'third'],
    ['a', 3, 'fourth'],
    ['a', 1, 'second'],
    ['b', 1, 'first'],
  ]);
  expect(result.restored).toEqual(result.sorted);
  expect(result.original).toEqual([['first'], ['second'], ['third'], ['fourth']]);
  expect(result.selection).toMatchObject({ startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 });
});

test('custom sort rejects duplicate columns and a stale dialog without writing', async ({
  page,
}) => {
  await workbookTool(page, 'sort-custom');
  await page.getByRole('button', { name: 'Add sort level' }).click();
  await page.getByLabel('Level 2 column').selectOption('0');
  await page.getByRole('button', { name: 'Apply sort' }).click();
  await expect(page.getByRole('alert')).toHaveText('Choose a different column for each level.');
  await page.getByLabel('Level 2 column').selectOption('1');
  const revision = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('D1')
      .setValues([['changed']]);
    return book.revision;
  });
  await page.getByRole('button', { name: 'Apply sort' }).click();
  await expect(page.getByRole('alert')).not.toHaveText('Sorting…');
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Close this dialog and select the rows again.',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog')).not.toBeVisible();
  await page.getByLabel('Read only', { exact: true }).check();
  await workbookTool(page, 'sort-custom', true);
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
});
