import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';

test('reports non-default selection and active tab without frozen panes', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'unfrozen-view.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: readFileSync('benchmarks/xlsx-unfrozen-view-openpyxl-2026-11-12.xlsx'),
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Worksheet tab selection is not imported');
  await expect(dialog).toContainText('Worksheet selection and active cell are not imported');
  await expect(dialog).not.toContainText('Some frozen or split worksheet panes are not imported');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'View', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await app.open(book.id);
    book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    return { freeze: book.sheetData(sheet.id).freeze, value: await sheet.range('A1').getValues() };
  });
  expect(result).toEqual({ freeze: { rows: 0, columns: 0 }, value: [[42]] });
});

test('imports frozen panes and reports split panes from an independent workbook', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  const upload = () =>
    page.locator('#file').setInputFiles({
      name: 'panes.xlsx',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      buffer: readFileSync('benchmarks/xlsx-pane-openpyxl-2026-11-12.xlsx'),
    });
  await upload();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Some frozen or split worksheet panes are not imported');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Frozen', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const initial = await book.getSheets()[0].range('A1').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    return {
      initial,
      values: await book.getSheets()[0].range('A1').getValues(),
      freeze: book.sheetData(book.getSheets()[0].id).freeze,
    };
  });
  expect(result).toEqual({ initial: [[7]], values: [[7]], freeze: { rows: 3, columns: 2 } });
});

test('imports three independent frozen-pane variants and retains them after reopening', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'freeze.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: readFileSync('benchmarks/xlsx-freeze-openpyxl-2026-11-12.xlsx'),
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Worksheet metadata !margins is not preserved');
  await expect(dialog).not.toContainText('Some frozen or split worksheet panes are not imported');
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Rows', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const get = () => book.getSheets().map((sheet) => book.sheetData(sheet.id).freeze);
    const before = get();
    await app.open(book.id);
    book = app.getWorkbook();
    return { before, after: get() };
  });
  const expected = [
    { rows: 4, columns: 0 },
    { rows: 0, columns: 3 },
    { rows: 3, columns: 2 },
  ];
  expect(result).toEqual({ before: expected, after: expected });
});

test('restores exported panes after a LibreOffice save and discloses view losses', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'freeze-libreoffice.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: readFileSync('benchmarks/xlsx-freeze-export-reader-2026-11-12.saved.xlsx'),
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).not.toContainText('Some frozen or split worksheet panes are not imported');
  await expect(dialog).toContainText('Worksheet tab selection is not imported');
  await expect(dialog).toContainText('Worksheet selection and active cell are not imported');
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Rows', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const snapshot = async () =>
      Promise.all(
        book.getSheets().map(async (sheet) => ({
          name: sheet.name,
          freeze: book.sheetData(sheet.id).freeze,
          value: await sheet.range('A1').getValues(),
        })),
      );
    const before = await snapshot();
    await app.open(book.id);
    book = app.getWorkbook();
    return { before, after: await snapshot() };
  });
  const expected = [
    { name: 'Rows', freeze: { rows: 4, columns: 0 }, value: [[100]] },
    { name: 'Columns', freeze: { rows: 0, columns: 3 }, value: [[101]] },
    { name: "Both '中' &", freeze: { rows: 3, columns: 2 }, value: [[102]] },
    { name: 'Unfrozen', freeze: { rows: 0, columns: 0 }, value: [[103]] },
  ];
  expect(result).toEqual({ before: expected, after: expected });
});
