import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';

test('identifies the correct sheets in export issues', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const sheets = window.opensheet.getWorkbook().getSheets();
    await sheets[0].range('A1').setValues([['_x0041_']]);
    await sheets[1].range('A1').setValues([['_x0042_']]);
  });
  await page.locator('#export-xlsx').click();
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(report.locator('li', { hasText: 'Literal _xHHHH_ sequences' })).toHaveCount(2);
  await expect(report.locator('li', { hasText: 'Launch budget: A1:' })).toBeVisible();
  await expect(report.locator('li', { hasText: 'Notes: A1:' })).toBeVisible();
});

test('identifies incoming sheet names in the import preview', async ({ page }) => {
  const workbook = XLSX.utils.book_new();
  for (const name of ['Incoming one', '<Incoming two>']) {
    const sheet = XLSX.utils.aoa_to_sheet([[1]]);
    sheet.A1.f = 'UNKNOWNFUNC(1)';
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  }
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'incoming.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(report.locator('li', { hasText: 'Incoming one: A1:' })).toContainText(
    'Formula uses a function',
  );
  await expect(report.locator('li', { hasText: '<Incoming two>: A1:' })).toContainText(
    'Formula uses a function',
  );
  await expect(report.locator('incoming')).toHaveCount(0);
  await expect(report.getByRole('button', { name: 'Import converted workbook' })).toBeVisible();
});

test('warns before exporting literal XLSX escape-shaped text', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1:A2')
      .setValues([['_x0041_'], ['before\0after']]);
  });
  await page.locator('#export-xlsx').click();
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toContainText(
    'Literal _xHHHH_ sequences may change when this XLSX is read',
  );
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toContainText(
    'Some spreadsheet applications may display control characters as escape sequences',
  );
});

test('warns before downloading XLSX with an unpaired surrogate', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['\uD800']]);
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText(
    'Unpaired UTF-16 surrogates may be replaced when written to XLSX',
  );
  await expect(dialog).toContainText('Launch budget: A1:');
  expect(downloads).toEqual([]);
});

test('warns that spreadsheet saves can remove link tooltips', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const range = window.opensheet.getWorkbook().getSheets()[0].range('A1');
    await range.setLink({ target: 'https://example.com/', tooltip: 'ordinary tip' });
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText(
    'Link tooltips may be removed when another spreadsheet application saves the file',
  );
  await expect(dialog).toContainText('Launch budget: A1:');
  expect(downloads).toEqual([]);
});

test('warns about long formulas before XLSX download', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A50')
      .setFormulas([['+1' + '+0'.repeat(4095)]]);
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Formulas of 8,192 or more UTF-16 code units');
  await expect(dialog).toContainText('Launch budget: A50:');
  expect(downloads).toEqual([]);
});

test('XLSX preview warns about blank-cell link round-trip risk', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const range = window.opensheet.getWorkbook().getSheets()[0].range('A1');
    await range.clear({ all: true });
    await range.setLink({ target: 'https://example.com/' });
  });
  await page.locator('#export-xlsx').click();
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toContainText(
    'Hyperlinks on blank cells may be lost when another spreadsheet application saves the file',
  );
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toContainText(
    'some readers may expose the target as the cell value',
  );
});

test('XLSX preview warns that empty text may be lost in another spreadsheet', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([['']]);
  });
  await page.locator('#export-xlsx').click();
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toContainText(
    'Empty text may be lost when another spreadsheet application saves the file',
  );
});

test('invalid retained compatibility report blocks XLSX download without changing the workbook', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const file = await book.toJSON();
    file.extensions['opensheet.importReport'] = {
      summary: { exact: 0, approximated: 0, dropped: 3, blocked: 0 },
    };
    await book.importJSON(file);
    if (!(await book.toJSON()).extensions['opensheet.importReport'])
      throw new Error('Imported report was lost');
    return {
      id: book.id,
      revision: book.revision,
      values: await book.getSheets()[0].range('A1:B2').getValues(),
    };
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  await expect(page.locator('#notice')).toContainText(
    'Invalid original import compatibility report',
  );
  await expect(page.getByRole('dialog', { name: 'Compatibility report' })).toBeHidden();
  expect(downloads).toEqual([]);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        id: book.id,
        revision: book.revision,
        values: await book.getSheets()[0].range('A1:B2').getValues(),
      };
    }),
  ).toEqual(before);
});

test('retained import losses remain visible after reopening before XLSX download', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const file = await book.toJSON();
    file.extensions['opensheet.importReport'] = {
      adapterVersion: '0.1.1',
      sourceFormat: 'SheetJS',
      summary: { exact: 0, approximated: 0, dropped: 3, blocked: 0 },
      issues: [
        {
          code: 'TEST_LOSS',
          severity: 'warning',
          feature: 'test',
          action: 'dropped',
          sheetId: 'removed_sheet',
          message: 'Original import omitted test features',
          count: 3,
        },
      ],
      omittedIssueOccurrences: 0,
    };
    await book.importJSON(file);
    await app.open(book.id);
  });
  const downloads: string[] = [];
  page.on('download', (download) => downloads.push(download.suggestedFilename()));
  await page.locator('#export-xlsx').click();
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText(
    'removed_sheet: Original import omitted test features (3 occurrences)',
  );
  expect(downloads).toEqual([]);
});
