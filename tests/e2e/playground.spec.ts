import { test, expect } from '@playwright/test';
import * as XLSX from 'xlsx';
import type { OpenSheet } from 'opensheet';
declare global {
  interface Window {
    opensheet: OpenSheet;
  }
}
test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.getByRole('grid', { name: 'Spreadsheet', exact: true })).toBeVisible();
  await expect(page.locator('#source')).toContainText('tabular');
});
test('renders the local workbook, formula totals and source preview', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await expect(page.getByRole('heading', { name: 'Your data, in good hands.' })).toBeVisible();
  const total = await page.evaluate(() =>
    window.opensheet.getWorkbook().getSheets()[0].range('F11').getValues(),
  );
  expect(total).toEqual([[13097]]);
  await page.getByRole('tab', { name: 'Markdown', exact: true }).click();
  await expect(page.locator('#source')).toContainText('| Item |');
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.screenshot({
    path: `test-results/playground-${test.info().project.name}.png`,
    fullPage: true,
  });
  expect(errors).toEqual([]);
});
test('edits with the DOM editor, recalculates, undoes and redoes', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).fill('D2');
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).press('Enter');
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.press('Enter');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('3');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('F2').getValues(),
    ),
  ).toEqual([[7200]]);
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('D2').getValues(),
    ),
  ).toEqual([[1]]);
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('D2').getValues(),
    ),
  ).toEqual([[3]]);
});
test('uses the formula bar, tab navigation and independent worksheets', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).fill('H2');
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).press('Enter');
  await page.getByRole('textbox', { name: 'Formula bar', exact: true }).fill('=SUM(D2:D10)');
  await page.getByRole('textbox', { name: 'Formula bar', exact: true }).press('Enter');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('H2').getValues(),
    ),
  ).toEqual([[15]]);
  await page.getByRole('tab', { name: 'Notes', exact: true }).click();
  await expect(page.getByRole('textbox', { name: 'Formula bar', exact: true })).toHaveValue(
    'OpenSheet',
  );
  await page.getByRole('button', { name: 'Add sheet', exact: true }).click();
  await expect(page.getByRole('tab', { name: 'Sheet1', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
});
test('blocks edits in read-only mode and preserves keyboard selection', async ({ page }) => {
  await page.getByLabel('Read only', { exact: true }).check();
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('x');
  await expect(page.getByRole('textbox', { name: 'Edit cell', exact: true })).toBeHidden();
  await grid.press('ArrowRight');
  await expect(page.getByRole('textbox', { name: 'Cell address', exact: true })).toHaveValue('B1');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['Item']]);
});
test('pastes a matrix atomically and protects IME composition from Enter', async ({ page }) => {
  await page.evaluate(() => {
    const app = window.opensheet,
      id = app.getWorkbook().getSheets()[0].id;
    app.selection.set({ sheetId: id, startRow: 15, endRow: 16, startColumn: 0, endColumn: 1 });
    app.getGrid()!.pasteText('姓名\t数量\n测试\t2');
  });
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A16:B17').getValues(),
    ),
  ).toEqual([
    ['姓名', '数量'],
    ['测试', 2],
  ]);
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.focus();
  await grid.press('Enter');
  const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
  await editor.dispatchEvent('compositionstart');
  await editor.fill('中文输入');
  await editor.dispatchEvent('keydown', { key: 'Enter', isComposing: true });
  await expect(editor).toBeVisible();
  await editor.dispatchEvent('compositionend');
  await editor.press('Enter');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A16').getValues(),
    ),
  ).toEqual([['中文输入']]);
});
test('imports a real XLSX in a Worker and exports a downloadable XLSX', async ({ page }) => {
  const wb = XLSX.utils.book_new(
    XLSX.utils.aoa_to_sheet([
      ['Name', 'Count'],
      ['Test', 7],
    ]),
    'Imported',
  );
  const bytes = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  await page.locator('#file').setInputFiles({
    name: 'sample.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  await expect(page.getByRole('tab', { name: 'Imported', exact: true })).toBeVisible();
  if (await page.getByRole('dialog').isVisible())
    await page.getByRole('button', { name: 'Close dialog', exact: true }).click();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual([[7]]);
  await page.getByRole('button', { name: 'Export XLSX', exact: false }).click();
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download converted XLSX', exact: true }).click();
  expect((await download).suggestedFilename()).toBe('sample.xlsx');
});
test('imports hostile JSON without replacing the current workbook', async ({ page }) => {
  await page.locator('#file').setInputFiles({
    name: 'bad.json',
    mimeType: 'application/json',
    buffer: Buffer.from('{"schemaVersion":99}'),
  });
  await expect(page.locator('#notice')).toContainText('Invalid');
  await expect(page.getByRole('tab', { name: 'Launch budget', exact: true })).toBeVisible();
});
test('provides an accessible data table and a functional plugin', async ({ page }) => {
  await page.evaluate(() => {
    const app = window.opensheet;
    app.selection.set({
      sheetId: app.getWorkbook().getSheets()[0].id,
      startRow: 1,
      endRow: 3,
      startColumn: 3,
      endColumn: 4,
    });
  });
  await page.getByRole('button', { name: 'Σ Sum', exact: true }).click();
  await expect(page.locator('.os-toast')).toContainText('sum to 2');
  await page.getByRole('button', { name: 'Table view', exact: true }).click();
  await expect(page.getByRole('cell', { name: 'Brand identity', exact: true })).toBeVisible();
});
test('loads saved JSON and cleans up old plugin UI', async ({ page }) => {
  const snapshot = await page.evaluate(() => window.opensheet.getWorkbook().toJSON());
  await page.locator('#file').setInputFiles({
    name: 'saved.json',
    mimeType: 'application/json',
    buffer: Buffer.from(JSON.stringify(snapshot)),
  });
  await expect(page.locator('#save-state')).toContainText('Imported locally');
  await expect(page.getByRole('button', { name: 'Σ Sum', exact: true })).toHaveCount(1);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('F11').getValues(),
    ),
  ).toEqual([[13097]]);
});
test('uses a responsive layout without body overflow', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await expect(page.getByRole('grid', { name: 'Spreadsheet', exact: true })).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  await page.evaluate(() => window.scrollTo(0, 0));
  await page.screenshot({
    path: `test-results/mobile-${test.info().project.name}.png`,
    fullPage: true,
  });
});

test('scrolls to distant selections and reads the edited frozen-grid cell', async ({ page }) => {
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).fill('J80');
  await page.getByRole('textbox', { name: 'Cell address', exact: true }).press('Enter');
  expect(await page.locator('.os-scroll').evaluate((e) => e.scrollTop)).toBeGreaterThan(1000);
  const grid = page.getByRole('grid', { name: 'Spreadsheet', exact: true });
  await grid.press('Enter');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('distant');
  await page.getByRole('textbox', { name: 'Edit cell', exact: true }).press('Enter');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('J80').getValues(),
    ),
  ).toEqual([['distant']]);
});
