import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';

test('warns about unsupported numeric display and retains the XLSX format code', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const workbook = {
    SheetNames: ['Data'],
    Sheets: { Data: { A1: { t: 'n', v: 1234, z: '0.00E+00' }, '!ref': 'A1' } },
  };
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  const upload = () =>
    page.locator('#file').setInputFiles({
      name: 'format.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: bytes,
    });
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await upload();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Number format is retained but may display approximately');
  await expect(page.getByRole('tab', { name: 'Data', exact: true })).toHaveCount(0);
  await dialog.getByRole('button', { name: 'Close dialog' }).click();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await upload();
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Data', exact: true })).toBeVisible();
  const detail = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const cell = await book.getCell(sheet.id, 0, 0);
    return { value: await sheet.range('A1').getValues(), format: cell?.numberFormat };
  });
  expect(detail).toEqual({ value: [[1234]], format: '0.00E+00' });
});

test('imports common slash date-time formats without losing the displayed time', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const workbook = {
    SheetNames: ['Dates'],
    Sheets: { Dates: { A1: { t: 'n', v: 45293.5, z: 'm/d/yy h:mm' }, '!ref': 'A1' } },
  };
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  await page.locator('#file').setInputFiles({
    name: 'dates.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  await expect(page.getByRole('tab', { name: 'Dates', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const before = await sheet.range('A1').getDisplayValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      before,
      after: await sheet.range('A1').getDisplayValues(),
      values: await sheet.range('A1').getValues(),
    };
  });
  expect(result).toEqual({
    before: [['1/2/24 12:00']],
    after: [['1/2/24 12:00']],
    values: [[45293.5]],
  });
});

test('imports and reopens negative currency with its sign before the currency symbol', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const workbook = {
    SheetNames: ['Currency'],
    Sheets: {
      Currency: {
        A1: { t: 'n', v: -1234.56, z: '$#,##0.00' },
        A2: { t: 'n', v: 1234.56, z: '$#,##0.00' },
        '!ref': 'A1:A2',
      },
    },
  };
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  await page.locator('#file').setInputFiles({
    name: 'currency.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  await expect(page.getByRole('tab', { name: 'Currency', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const before = await sheet.range('A1:A2').getDisplayValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      before,
      after: await sheet.range('A1:A2').getDisplayValues(),
      values: await sheet.range('A1:A2').getValues(),
      format: (await book.getCell(sheet.id, 0, 0))?.numberFormat,
    };
  });
  expect(result).toEqual({
    before: [['-$1,234.56'], ['$1,234.56']],
    after: [['-$1,234.56'], ['$1,234.56']],
    values: [[-1234.56], [1234.56]],
    format: '$#,##0.00',
  });
});
