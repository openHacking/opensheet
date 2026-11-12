import { expect, test } from '@playwright/test';

test('find reports incomplete calculation and keeps the selection', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1').setFormulas([['#LIMIT!']]);
    await sheet.range('B1').setValues([['needle']]);
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 });
  });
  await page.locator('#find').fill('needle');
  await page.locator('#find-next').click();
  await expect(page.locator('#notice')).toContainText(
    'Search calculation exceeded its work budget',
  );
  expect(await page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(1);
  await page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].range('A1').clear());
  await page.locator('#find-next').click();
  await expect(page.locator('#notice')).toContainText('1 matching cells');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.selection.get()?.startColumn))
    .toBe(1);
});

test('search rejects exhausted calculation without caching partial results', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = await book.addSheet('Search limit', { rows: 5, columns: 2 });
    await sheet.range('A1').setFormulas([['#LIMIT!']]);
    await sheet.range('B1').setValues([['#LIMIT!']]);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheetByName('Search limit')!;
    const codes: string[] = [];
    for (const direction of ['next', 'previous'] as const)
      codes.push(
        await book.search(sheet.id, '#LIMIT!', -1, { direction }).then(
          () => 'success',
          (error: { code: string }) => error.code,
        ),
      );
    await sheet.range('A1').clear();
    const complete = await book.search(sheet.id, '#LIMIT!');
    await book.undo();
    const afterUndo = await book.search(sheet.id, '#LIMIT!').then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    return { codes, complete, afterUndo };
  });
  expect(result).toEqual({
    codes: ['LIMIT_EXCEEDED', 'LIMIT_EXCEEDED'],
    complete: { count: 1, match: { row: 0, column: 1 } },
    afterUndo: 'LIMIT_EXCEEDED',
  });
});
