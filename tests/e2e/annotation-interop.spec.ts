import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { FileReader } from '@opensheetjs/core';

test('imports LibreOffice blank-cell notes through the file Worker and exports them again', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await expect(page.locator('#app')).not.toHaveAttribute('inert', '');
  await page.locator('#file').setInputFiles({
    name: 'annotations.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: readFileSync(
      new URL('../../benchmarks/annotation-interop-2026-10-31.reader.xlsx', import.meta.url),
    ),
  });
  await page.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Annotations', exact: true })).toBeVisible();
  if (await page.getByRole('dialog').isVisible())
    await page.getByRole('button', { name: 'Close dialog' }).click();
  const state = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const initial = {
      notes: await sheet.range('A1:A4').getNotes(),
      values: await sheet.range('A1:A4').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = {
      notes: await sheet.range('A1:A4').getNotes(),
      values: await sheet.range('A1:A4').getValues(),
    };
    return { initial, reopened };
  });
  const expected = {
    notes: [[''], [''], ['hello'], ['hello']],
    values: [['value'], [null], ['value'], [null]],
  };
  expect(state).toEqual({ initial: expected, reopened: expected });
  await page.getByRole('button', { name: 'Export XLSX', exact: false }).click();
  const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download converted XLSX' }).click();
  const download = await waiting,
    stream = await download.createReadStream();
  const chunks: Buffer[] = [];
  for await (const chunk of stream) chunks.push(Buffer.from(chunk));
  const file = fromSheetJS(
    XLSX.read(Buffer.concat(chunks), {
      type: 'buffer',
      dense: true,
      cellStyles: true,
      cellNF: true,
    }),
  ).file;
  const reader = new FileReader(file);
  try {
    const sheetId = file.sheets[0].id;
    expect([0, 1, 2, 3].map((row) => reader.getCell(sheetId, row, 0)?.note ?? null)).toEqual([
      '',
      '',
      'hello',
      'hello',
    ]);
    expect([0, 1, 2, 3].map((row) => reader.value(sheetId, row, 0))).toEqual([
      'value',
      null,
      'value',
      null,
    ]);
  } finally {
    reader.dispose();
  }
});
