import { expect, test } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import * as XLSX from 'xlsx';

test('reports oversized XLSX text, preserves the workbook and exports after correction', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const revision = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['x'.repeat(32768)]]);
    return book.revision;
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(report).toContainText('32,767 UTF-16 code units');
  await expect(report).toContainText('First location: Launch budget!A1');
  await expect(report).toContainText('A1:');
  await expect(report.getByRole('button', { name: 'Download converted XLSX' })).toHaveCount(0);
  expect(downloads).toEqual([]);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        revision: book.revision,
        value: (await book.getSheets()[0].range('A1').getValues())[0][0],
      };
    }),
  ).toEqual({ revision, value: 'x'.repeat(32768) });
  await page.keyboard.press('Escape');
  await page.evaluate(async () => {
    const app = window.opensheet;
    await app.open(app.getWorkbook().id);
    const sheet = app.getWorkbook().getSheets()[0];
    if ((await sheet.range('A1').getValues())[0][0] !== 'x'.repeat(32768))
      throw new Error('value changed');
    await sheet.range('A1').setValues([['x'.repeat(32767)]]);
  });
  const download = page.waitForEvent('download');
  await page.locator('#export-xlsx').click();
  const preview = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(preview).toBeVisible();
  await preview.getByRole('button', { name: 'Download converted XLSX' }).click();
  const path = await (await download).path();
  expect(path).toBeTruthy();
  const file = XLSX.read(await readFile(path!), { type: 'buffer' });
  expect(file.Sheets[file.SheetNames[0]].A1.v).toBe('x'.repeat(32767));
});
