import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';
import { readFileSync } from 'node:fs';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';

test('exports Unicode links with a warning and reopens equivalent destinations', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const targets = [
    'https://例子.测试/é/🙂?x=1&y=2#片段',
    'mailto:用户@example.com?subject=你好',
    'https://example.com/a%20b?x=1&y=2#frag',
  ];
  await page.evaluate(async (values) => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    await sheet.range('A1:A3').setValues([['one'], ['two'], ['three']]);
    for (let i = 0; i < values.length; i++)
      await sheet.range(`A${i + 1}`).setLink({ target: values[i] });
  }, targets);
  await page.locator('#export-xlsx').click();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Unicode link targets are normalized to URL-encoded ASCII');
  await expect(dialog).toContainText('Launch budget: A1:');
  const pending = page.waitForEvent('download');
  await dialog.getByRole('button', { name: 'Download converted XLSX' }).click();
  const bytes = readFileSync((await (await pending).path())!);
  const parsed = XLSX.read(bytes, { type: 'buffer', cellStyles: true, cellNF: true });
  const converted = fromSheetJS(parsed);
  const expected = targets.map((target) => new URL(target).href);
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    await app.load(JSON.parse(fileJSON));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const initial = (await sheet.range('A1:A3').getLinks()).map((row) => row[0]?.target);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = (await sheet.range('A1:A3').getLinks()).map((row) => row[0]?.target);
    return { initial, reopened };
  }, JSON.stringify(converted.file));
  expect(result).toEqual({ initial: expected, reopened: expected });
});

test('restores Unicode links from an independently saved workbook', async ({ page }) => {
  const bytes = readFileSync('benchmarks/xlsx-link-reader-2026-11-21.saved.xlsx');
  const evidence = JSON.parse(readFileSync('benchmarks/xlsx-link-reader-2026-11-21.json', 'utf8'));
  const expected = evidence.savedInspection.links
    .slice(0, 2)
    .map(
      (link: { target: string; location: string | null }) =>
        link.target + (link.location === null ? '' : `#${link.location}`),
    );
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'links.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText(
    'Unicode link target restored from a verified XLSX relationship',
  );
  await expect(dialog).toContainText('Links: A1:');
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Links', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const before = (await sheet.range('A1:A2').getLinks()).map((row) => row[0]?.target);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { before, after: (await sheet.range('A1:A2').getLinks()).map((row) => row[0]?.target) };
  });
  expect(result).toEqual({ before: expected, after: expected });
});
