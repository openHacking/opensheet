import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';

test('compressed XLSX preflight imports and reopens all fixture values', async ({ page }) => {
  const rows = Array.from({ length: 128 }, (_, i) => [i, `row-${i}-${'x'.repeat(80)}`]);
  const workbook = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet(rows), 'Compressed');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', compression: true });
  let deflated = 0;
  for (let offset = 0; offset + 46 <= buffer.length; offset++)
    if (buffer.readUInt32LE(offset) === 0x02014b50 && buffer.readUInt16LE(offset + 10) === 8)
      deflated++;
  expect(deflated).toBeGreaterThan(0);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'compressed.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer,
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Compressed', exact: true })).toBeVisible();
  const actual = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const first = await book.getSheets()[0].range('A1:B128').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    return { first, reopened: await book.getSheets()[0].range('A1:B128').getValues() };
  });
  expect(actual.first).toEqual(rows);
  expect(actual.reopened).toEqual(rows);
});
