import { expect, test } from '@playwright/test';

test('Control+End navigates to the used extent and Shift extends from the keyboard anchor', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 100, columns: 20 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('F40').setValues([[42]]);
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const grid = page.getByRole('grid', { exact: true });
  const selection = () => page.evaluate(() => window.opensheet.selection.get());
  await grid.focus();
  await page.keyboard.press('Control+Shift+End');
  expect(await selection()).toMatchObject({
    startRow: 1,
    endRow: 40,
    startColumn: 1,
    endColumn: 6,
  });
  await page.keyboard.press('Shift+ArrowLeft');
  expect(await selection()).toMatchObject({
    startRow: 1,
    endRow: 40,
    startColumn: 1,
    endColumn: 5,
  });
  await page.keyboard.press('Control+End');
  expect(await selection()).toMatchObject({
    startRow: 39,
    endRow: 40,
    startColumn: 5,
    endColumn: 6,
  });
  expect(await page.evaluate(() => window.opensheet.getGrid()!.scroller.scrollTop)).toBeGreaterThan(
    0,
  );
});

test('Control+End resolves hidden used edges and remains in a cell draft while editing', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 50, columns: 10 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('F20').setValues([[42]]);
    await sheet.setRowsHidden(19, 1, true);
    await sheet.setColumnsHidden(5, 1, true);
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await page.keyboard.press('Control+End');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 20,
    endRow: 21,
    startColumn: 6,
    endColumn: 7,
  });
  await page.keyboard.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await editor.fill('draft');
  await editor.press('Control+End');
  await expect(editor).toHaveValue('draft');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 20,
    startColumn: 6,
  });
});

test('Meta+End handles an empty sheet and expands the destination merge', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 20, columns: 10 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 4, endRow: 5, startColumn: 4, endColumn: 5 });
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await page.keyboard.press('Meta+End');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 0,
    endRow: 1,
    startColumn: 0,
    endColumn: 1,
  });
  await page.evaluate(async () => {
    await window.opensheet.getWorkbook().getSheets()[0].range('E10:F11').merge();
  });
  await grid.focus();
  await page.keyboard.press('Meta+End');
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 9,
    endRow: 11,
    startColumn: 4,
    endColumn: 6,
  });
});
