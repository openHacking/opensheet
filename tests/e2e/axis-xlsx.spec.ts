import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';

test('imports independent axis dimensions and warns before XLSX export', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'dimensions.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: readFileSync('benchmarks/axis-dimension-openpyxl-2026-11-12.xlsx'),
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Spreadsheet column width is font-dependent');
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Dimensions', exact: true })).toBeVisible();
  const metadata = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    await app.open(book.id);
    const restored = app.getWorkbook();
    const sheet = restored.sheetData(restored.getSheets()[0].id);
    return { row: sheet.rows[sheet.rowOrder[100]], column: sheet.columns[sheet.columnOrder[20]] };
  });
  expect(metadata.row).toEqual({ size: 42, hidden: true });
  expect(metadata.column.hidden).toBe(true);
  await page.locator('#export-xlsx').click();
  await expect(dialog).toContainText(
    'Pixel column widths are converted to font-dependent spreadsheet widths',
  );
  const downloaded = page.waitForEvent('download');
  await dialog.getByRole('button', { name: 'Download converted XLSX' }).click();
  const file = await downloaded;
  const bytes = readFileSync((await file.path())!);
  const output = XLSX.read(bytes, { type: 'buffer', cellStyles: true });
  expect(output.Sheets.Dimensions['!rows']![100]).toMatchObject({ hpt: 31.5, hidden: true });
  expect(output.Sheets.Dimensions['!cols']![20]).toMatchObject({ hidden: true });
});
