import { expect, test } from '@playwright/test';

test('keyboard selection expands, shrinks and crosses its original anchor', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const select = async (row: number, column: number) => {
    await page.evaluate(
      ({ row, column }) => {
        const app = window.opensheet;
        app.selection.set({
          sheetId: app.getWorkbook().getSheets()[0].id,
          startRow: row,
          endRow: row + 1,
          startColumn: column,
          endColumn: column + 1,
        });
      },
      { row, column },
    );
    await page.getByRole('grid', { exact: true }).focus();
  };
  const bounds = async () =>
    page.evaluate(() => {
      const { startRow, endRow, startColumn, endColumn } = window.opensheet.selection.get()!;
      return { startRow, endRow, startColumn, endColumn };
    });
  await select(4, 4);
  await page.keyboard.press('Shift+ArrowUp');
  await page.keyboard.press('Shift+ArrowUp');
  expect(await bounds()).toEqual({ startRow: 2, endRow: 5, startColumn: 4, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 3, endRow: 5, startColumn: 4, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 4, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 6, startColumn: 4, endColumn: 5 });
  await select(7, 7);
  await page.keyboard.press('Shift+ArrowLeft');
  await page.keyboard.press('Shift+ArrowLeft');
  expect(await bounds()).toEqual({ startRow: 7, endRow: 8, startColumn: 5, endColumn: 8 });
  await page.keyboard.press('Shift+ArrowRight');
  expect(await bounds()).toEqual({ startRow: 7, endRow: 8, startColumn: 6, endColumn: 8 });
  await page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].setColumnHidden(3, true));
  await select(4, 4);
  await page.keyboard.press('Shift+ArrowLeft');
  await page.keyboard.press('Shift+ArrowLeft');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 1, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowRight');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 2, endColumn: 5 });
});

test('keyboard extension skips hidden rows and traverses merged cells', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.setRowHidden(3, true);
    await sheet.range('C7:D8').clear({ all: true });
    await sheet.range('C7:D8').merge();
    app.selection.set({ sheetId: sheet.id, startRow: 4, endRow: 5, startColumn: 4, endColumn: 5 });
  });
  await page.getByRole('grid', { exact: true }).focus();
  const bounds = async () =>
    page.evaluate(() => {
      const { startRow, endRow, startColumn, endColumn } = window.opensheet.selection.get()!;
      return { startRow, endRow, startColumn, endColumn };
    });
  await page.keyboard.press('Shift+ArrowUp');
  await page.keyboard.press('Shift+ArrowUp');
  expect(await bounds()).toEqual({ startRow: 1, endRow: 5, startColumn: 4, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 2, endRow: 5, startColumn: 4, endColumn: 5 });
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 6,
      endRow: 7,
      startColumn: 2,
      endColumn: 3,
    });
  });
  await page.keyboard.press('Shift+ArrowUp');
  await page.keyboard.press('Shift+ArrowUp');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 8, startColumn: 2, endColumn: 4 });
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 6, endRow: 8, startColumn: 2, endColumn: 4 });
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 6, endRow: 9, startColumn: 2, endColumn: 4 });
  await page.keyboard.press('ArrowDown');
  expect(await bounds()).toEqual({ startRow: 6, endRow: 8, startColumn: 2, endColumn: 4 });
});

test('Shift+Home extends to row or worksheet origin without losing the anchor', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
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
  await page.getByRole('grid', { exact: true }).focus();
  const bounds = async () =>
    page.evaluate(() => {
      const { startRow, endRow, startColumn, endColumn } = window.opensheet.selection.get()!;
      return { startRow, endRow, startColumn, endColumn };
    });
  await page.keyboard.press('Shift+Home');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 0, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowRight');
  expect(await bounds()).toEqual({ startRow: 4, endRow: 5, startColumn: 1, endColumn: 5 });
  await page.keyboard.press('Control+Shift+Home');
  expect(await bounds()).toEqual({ startRow: 0, endRow: 5, startColumn: 0, endColumn: 5 });
  await page.keyboard.press('Shift+ArrowDown');
  expect(await bounds()).toEqual({ startRow: 1, endRow: 5, startColumn: 0, endColumn: 5 });
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 7,
      endRow: 8,
      startColumn: 7,
      endColumn: 8,
    });
  });
  await page.keyboard.press('Meta+Shift+Home');
  expect(await bounds()).toEqual({ startRow: 0, endRow: 8, startColumn: 0, endColumn: 8 });
});
