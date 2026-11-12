import { expect, test } from '@playwright/test';

test('keyboard movement remains visible when worksheet edges are hidden', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Small', rows: 6, columns: 6 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setRowHidden(0, true);
    await sheet.setRowHidden(5, true);
    await sheet.setColumnHidden(0, true);
    await sheet.setColumnHidden(5, true);
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  await page.getByRole('grid', { exact: true }).focus();
  const bounds = () =>
    page.evaluate(() => {
      const { startRow, endRow, startColumn, endColumn } = window.opensheet.selection.get()!;
      return { startRow, endRow, startColumn, endColumn };
    });
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('ArrowLeft');
  expect(await bounds()).toEqual({ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 4,
      endRow: 5,
      startColumn: 4,
      endColumn: 5,
    });
  });
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 4, endColumn: 5 });
  await page.keyboard.press('Home');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 1, endColumn: 2 });
  await page.keyboard.press('Control+Home');
  expect(await bounds()).toEqual({ startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 4,
      endRow: 5,
      startColumn: 4,
      endColumn: 5,
    });
  });
  await page.keyboard.press('Control+Shift+Home');
  expect(await bounds()).toEqual({ startRow: 1, endRow: 5, startColumn: 1, endColumn: 5 });
});
