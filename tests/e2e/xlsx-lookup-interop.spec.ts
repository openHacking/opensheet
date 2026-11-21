import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

const source = readFileSync(
  new URL('../../benchmarks/lookup-error-reader-2026-11-22.source.xlsx', import.meta.url),
);
const warning =
  'MATCH and VLOOKUP may handle lookup errors, blank cells or wildcards differently across spreadsheet applications';

test('previews lookup semantics before XLSX import and export', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => ({
    id: window.opensheet.getWorkbook().id,
    revision: window.opensheet.getWorkbook().revision,
  }));
  await page
    .locator('#file')
    .setInputFiles({ name: 'lookup.xlsx', mimeType: 'application/octet-stream', buffer: source });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText(warning);
  expect(
    await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    })),
  ).toEqual(before);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Lookup', exact: true })).toBeVisible();
  const values = await page.evaluate(async () => {
    const app = window.opensheet;
    const initial = await app.getWorkbook().getSheets()[0].range('D1:D3').getValues();
    await app.open(app.getWorkbook().id);
    return { initial, reopened: await app.getWorkbook().getSheets()[0].range('D1:D3').getValues() };
  });
  expect(values).toEqual({
    initial: [[{ error: '#DIV/0!' }], [1], [{ error: '#DIV/0!' }]],
    reopened: [[{ error: '#DIV/0!' }], [1], [{ error: '#DIV/0!' }]],
  });
  await page.getByRole('button', { name: 'Export XLSX', exact: false }).click();
  const exportDialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(exportDialog).toContainText(warning);
  const download = page.waitForEvent('download');
  await exportDialog.getByRole('button', { name: 'Download converted XLSX' }).click();
  expect((await download).suggestedFilename()).toMatch(/\.xlsx$/);
});
