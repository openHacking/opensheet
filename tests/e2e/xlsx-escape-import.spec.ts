import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { xlsxEntityEscape } from '../fixtures/xlsx-entity-escape.js';

test('warns before importing a workbook with ambiguous XLSX escapes', async ({ page }) => {
  const bytes = readFileSync(
    new URL('../../benchmarks/xlsx-escape-reader-2026-11-12.xlsx', import.meta.url),
  );
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'escape.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('XLSX escape-shaped text may have changed during parsing');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'A', exact: true })).toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['A']]);
});

test('warns about escape-shaped text formed by XML character references', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'entity.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: xlsxEntityEscape(),
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('XLSX escape-shaped text may have changed during parsing');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Data', exact: true })).toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['A']]);
});
