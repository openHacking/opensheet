import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '../../packages/adapter-sheetjs/src/index.js';

const buffer = readFileSync(
  new URL('../../benchmarks/xlsx-graphics.openpyxl.xlsx', import.meta.url),
);

test('reports independently authored chart and media before replacing the workbook', async ({
  page,
}) => {
  const workbook = XLSX.read(buffer, { bookFiles: true });
  expect(
    fromSheetJS(workbook).report.issues.filter((issue) => issue.feature.endsWith('-parts')),
  ).toEqual([
    expect.objectContaining({ code: 'CHART_PARTS_DROPPED', count: 1, action: 'dropped' }),
    expect.objectContaining({ code: 'MEDIA_PARTS_DROPPED', count: 1, action: 'dropped' }),
  ]);
  expect(() => fromSheetJS(workbook, { unsupported: 'strict' })).toThrow(
    'Strict compatibility check failed',
  );
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => ({
    id: window.opensheet.getWorkbook().id,
    revision: window.opensheet.getWorkbook().revision,
  }));
  await page
    .locator('#file')
    .setInputFiles({ name: 'graphics.xlsx', mimeType: 'application/octet-stream', buffer });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Chart parts in the XLSX package are not imported');
  await expect(dialog).toContainText('Embedded media files in the XLSX package are not imported');
  expect(
    await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    })),
  ).toEqual(original);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Graphics data', exact: true })).toBeVisible();
  const values = await page.evaluate(async () => {
    const app = window.opensheet;
    const first = await app.getWorkbook().getSheets()[0].range('A1:B2').getValues();
    await app.open(app.getWorkbook().id);
    return { first, reopened: await app.getWorkbook().getSheets()[0].range('A1:B2').getValues() };
  });
  expect(values).toEqual({
    first: [
      ['Sales', 5],
      ['Costs', 3],
    ],
    reopened: [
      ['Sales', 5],
      ['Costs', 3],
    ],
  });
});

test('only reports conventional files, and strict mode rejects each observed kind', () => {
  const base = { SheetNames: ['Sheet'], Sheets: { Sheet: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } } };
  const part = { content: new Uint8Array([1]) };
  expect(() => fromSheetJS(base, { unsupported: 'strict' })).not.toThrow();
  const ignored = {
    ...base,
    files: { 'xl/charts/style1.xml': part, 'xl/charts/chart1.xml.rels': part, 'xl/media/': part },
  };
  expect(fromSheetJS(ignored).report.issues.some((issue) => issue.feature.endsWith('-parts'))).toBe(
    false,
  );
  for (const path of ['XL/CHARTS/CHART1.XML', 'XL/MEDIA/image1.png']) {
    const input = { ...base, files: { [path]: part } };
    expect(fromSheetJS(input).report.issues.some((issue) => issue.feature.endsWith('-parts'))).toBe(
      true,
    );
    expect(() => fromSheetJS(input, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  }
});
