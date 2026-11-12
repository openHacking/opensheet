import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';

test('imports delimited text with empty fields after enabling worksheet stubs', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'blank-fields.csv',
    mimeType: 'text/csv',
    buffer: Buffer.from('left,,right\n1,,3'),
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toBeVisible();
  const confirm = dialog.getByRole('button', { name: 'Import converted workbook' });
  if (await confirm.isVisible()) await confirm.click();
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().id))
    .not.toBe(original);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:C2').getValues(),
    ),
  ).toEqual([
    ['left', null, 'right'],
    [1, null, 3],
  ]);
});

test('imports independent formulas without cached values and recalculates after reopen', async ({
  page,
}) => {
  const buffer = readFileSync(
    new URL('../../benchmarks/formula-uncached-openpyxl-2026-11-11.xlsx', import.meta.url),
  );
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page
    .locator('#file')
    .setInputFiles({ name: 'uncached.xlsx', mimeType: 'application/octet-stream', buffer });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Formula uses a function that OpenSheet does not support');
  const confirm = dialog.getByRole('button', { name: 'Import converted workbook' });
  if (await confirm.isVisible()) await confirm.click();
  await expect(page.getByRole('tab', { name: 'Sheet', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const first = await sheet.range('A1:A3').getValues();
    const cached = await Promise.all(
      [0, 1, 2].map(async (row) => !!(await book.getCell(sheet.id, row, 0))?.cached),
    );
    const display = await sheet.range('A3').getDisplayValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { first, cached, display, reopened: await sheet.range('A1:A3').getValues() };
  });
  expect(result).toEqual({
    first: [[64], [1], [{ error: '#NAME?' }]],
    cached: [false, false, false],
    display: [['#NAME?']],
    reopened: [[64], [1], [{ error: '#NAME?' }]],
  });
});

test('imports supported exponent formulas with source precedence and reopens', async ({ page }) => {
  const sheet = XLSX.utils.aoa_to_sheet([[null], [null]]);
  sheet.A1 = { t: 'n', f: '2^3^2', v: 64 };
  sheet.A2 = { t: 'n', f: '-2^2', v: 4 };
  sheet.A3 = { t: 'n', f: '2^+3^2', v: 64 };
  sheet['!ref'] = 'A1:A3';
  const buffer = XLSX.write(XLSX.utils.book_new(sheet, 'Power compatible'), {
    type: 'buffer',
    bookType: 'xlsx',
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page
    .locator('#file')
    .setInputFiles({ name: 'power.xlsx', mimeType: 'application/octet-stream', buffer });
  await expect(page.getByRole('tab', { name: 'Power compatible', exact: true })).toBeVisible();
  const values = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const first = await book.getSheets()[0].range('A1:A3').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    return { first, reopened: await book.getSheets()[0].range('A1:A3').getValues() };
  });
  expect(values).toEqual({ first: [[64], [4], [64]], reopened: [[64], [4], [64]] });
});

test('warns about differing exponent precedence before importing XLSX', async ({ page }) => {
  const sheet = XLSX.utils.aoa_to_sheet([[null], [null]]);
  sheet.A1 = { t: 'n', f: '2^3^2', v: 64 };
  sheet.A2 = { t: 'n', f: '-2^2', v: 4 };
  sheet.A3 = { t: 'n', f: 'UNSUPPORTED(1)^-2', v: 0 };
  sheet['!ref'] = 'A1:A3';
  const workbook = XLSX.utils.book_new(sheet, 'Power precedence');
  const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page
    .locator('#file')
    .setInputFiles({ name: 'precedence.xlsx', mimeType: 'application/octet-stream', buffer });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Exponentiation and negation may calculate differently');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Power precedence', exact: true })).toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    ),
  ).toEqual([[64], [4]]);
});
import { readFileSync } from 'node:fs';
import { fromSheetJS } from '../../packages/adapter-sheetjs/src/index.js';

test('rejects unterminated worksheet markup without replacing the workbook', async ({ page }) => {
  const source = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([[5]]), 'Malformed');
  const zip = XLSX.CFB.read(XLSX.write(source, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
  });
  const part = XLSX.CFB.find(zip, '/xl/worksheets/sheet1.xml');
  part.content = Buffer.concat([Buffer.from(part.content), Buffer.from('<!--'.repeat(20_000))]);
  part.size = part.content.length;
  const buffer = XLSX.CFB.write(zip, {
    type: 'buffer',
    fileType: 'zip',
    compression: true,
  }) as Buffer;
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => ({
    id: window.opensheet.getWorkbook().id,
    revision: window.opensheet.getWorkbook().revision,
  }));
  await page
    .locator('#file')
    .setInputFiles({ name: 'malformed.xlsx', mimeType: 'application/octet-stream', buffer });
  await expect(page.locator('#notice')).toContainText('Unterminated XLSX XML markup');
  expect(
    await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    })),
  ).toEqual(before);
});

test('rejects an oversized XLSX formula without replacing the current workbook', async ({
  page,
}) => {
  const sheet = XLSX.utils.aoa_to_sheet([[1]]);
  sheet.A1 = { t: 'n', f: 'A'.repeat(32_769), v: 1 };
  const book = XLSX.utils.book_new(sheet, 'Oversized');
  const buffer = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const current = window.opensheet.getWorkbook();
    return {
      id: current.id,
      revision: current.revision,
      cells: await current.getSheets()[0].range('A1:B3').getInputs(),
    };
  });
  await page
    .locator('#file')
    .setInputFiles({ name: 'oversized.xlsx', mimeType: 'application/octet-stream', buffer });
  await expect(page.locator('#notice')).toContainText('Formula exceeds text limit');
  expect(
    await page.evaluate(async () => {
      const app = window.opensheet;
      await app.open(app.getWorkbook().id);
      const current = app.getWorkbook();
      return {
        id: current.id,
        revision: current.revision,
        cells: await current.getSheets()[0].range('A1:B3').getInputs(),
      };
    }),
  ).toEqual(before);
});

test('preserves long formula inputs across XLSX import and reopen', async ({ page }) => {
  const formulas = ['A'.repeat(32_000), '['.repeat(32_000), 'SUM(1,UNKNOWN(2))'];
  const sheet = XLSX.utils.aoa_to_sheet([[1], [2], [3]]);
  for (let row = 0; row < formulas.length; row++)
    sheet[`A${row + 1}`] = { t: 'n', f: formulas[row], v: row + 1 };
  const book = XLSX.utils.book_new(sheet, 'Long formulas');
  const buffer = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page
    .locator('#file')
    .setInputFiles({ name: 'long-formulas.xlsx', mimeType: 'application/octet-stream', buffer });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Formula uses a function that OpenSheet does not support');
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Long formulas', exact: true })).toBeVisible();
  if (await page.getByRole('dialog').isVisible())
    await page.getByRole('button', { name: 'Close dialog' }).click();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const read = async () => {
      const current = app.getWorkbook();
      const sheet = current.getSheets()[0];
      return Promise.all(
        [0, 1, 2].map(async (row) => (await current.getCell(sheet.id, row, 0))?.input),
      );
    };
    const initial = await read();
    await app.open(app.getWorkbook().id);
    return { initial, reopened: await read() };
  });
  const expected = formulas.map((expression) => ({ type: 'formula', expression }));
  expect(result).toEqual({ initial: expected, reopened: expected });
});

test('warns about an unsupported function before importing XLSX', async ({ page }) => {
  const sheet = XLSX.utils.aoa_to_sheet([[1]]);
  sheet.A1 = { t: 'n', f: 'UNKNOWN(1)', v: 42 };
  const book = XLSX.utils.book_new(sheet, 'Functions');
  const buffer = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page
    .locator('#file')
    .setInputFiles({ name: 'function.xlsx', mimeType: 'application/octet-stream', buffer });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText('Formula uses a function that OpenSheet does not support');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
});

test('discloses external formula references before import without fetching them', async ({
  page,
}) => {
  const sheet = XLSX.utils.aoa_to_sheet([[1]]);
  sheet.A1 = { t: 'n', f: "'[external.invalid]Data'!A1", v: 42 };
  const book = XLSX.utils.book_new(sheet, 'External');
  const buffer = Buffer.from(XLSX.write(book, { type: 'buffer', bookType: 'xlsx' }));
  const requests: string[] = [];
  page.on('request', (request) => {
    if (request.url().includes('external.invalid')) requests.push(request.url());
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page
    .locator('#file')
    .setInputFiles({ name: 'external.xlsx', mimeType: 'application/octet-stream', buffer });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toContainText(
    'External workbook and structured table references are not evaluated',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'External', exact: true })).toBeVisible();
  expect(requests).toEqual([]);
});

function fixture(relocated: boolean, mixedCase = false) {
  const source = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([[5]]), 'Rule data');
  const zip = XLSX.CFB.read(XLSX.write(source, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
  });
  const path = 'xl/worksheets/sheet1.xml';
  const part = XLSX.CFB.find(zip, `/${path}`);
  const xml = Buffer.from(part.content).toString('utf8');
  const rules = `<!-- <dataValidation/> <conditionalFormatting/> -->
  <conditionalFormatting sqref="A1"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>
  <dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A1"><formula1>0</formula1></dataValidation></dataValidations>`;
  part.content = Buffer.from(xml.replace('</worksheet>', `${rules}</worksheet>`));
  part.size = part.content.length;
  if (relocated) {
    const replacement = 'xl/custom/rules.xml';
    const index = zip.FileIndex.indexOf(part);
    zip.FullPaths[index] = `Root Entry/${replacement}`;
    part.name = 'rules.xml';
    for (const name of ['/[Content_Types].xml', '/xl/_rels/workbook.xml.rels']) {
      const item = XLSX.CFB.find(zip, name);
      const text = Buffer.from(item.content).toString('utf8');
      item.content = Buffer.from(text.replaceAll('worksheets/sheet1.xml', 'custom/rules.xml'));
      item.size = item.content.length;
    }
  }
  if (mixedCase) {
    const manifest = XLSX.CFB.find(zip, '/[Content_Types].xml');
    manifest.content = Buffer.from(
      Buffer.from(manifest.content)
        .toString('utf8')
        .replaceAll('worksheets/sheet1.xml', 'worksheets/SHEET1.XML'),
    );
    manifest.size = manifest.content.length;
    const index = zip.FileIndex.indexOf(manifest);
    zip.FullPaths[index] = 'Root Entry/[content_types].xml';
    manifest.name = '[content_types].xml';
  }
  return XLSX.CFB.write(zip, { type: 'buffer', fileType: 'zip', compression: true }) as Buffer;
}

for (const source of ['conventional', 'declared custom', 'mixed case', 'openpyxl'] as const) {
  test(`discloses dropped worksheet rules in ${source} XLSX`, async ({ page }) => {
    const independent = source === 'openpyxl';
    const buffer = independent
      ? readFileSync(new URL('../../benchmarks/xlsx-rule-loss.openpyxl.xlsx', import.meta.url))
      : fixture(source === 'declared custom', source === 'mixed case');
    const workbook = XLSX.read(buffer, { bookFiles: true });
    const report = fromSheetJS(workbook).report;
    expect(
      report.issues.filter((issue) =>
        ['data-validation', 'conditional-formatting'].includes(issue.feature),
      ),
    ).toEqual([
      expect.objectContaining({ feature: 'data-validation', count: independent ? 3 : 1 }),
      expect.objectContaining({ feature: 'conditional-formatting', count: independent ? 2 : 1 }),
    ]);
    expect(() => fromSheetJS(workbook, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const original = await page.evaluate(() => window.opensheet.getWorkbook().id);
    await page
      .locator('#file')
      .setInputFiles({ name: 'rules.xlsx', mimeType: 'application/octet-stream', buffer });
    const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
    await expect(dialog).toContainText('Data validation rules are not imported');
    await expect(dialog).toContainText('Conditional formatting is not imported');
    expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
    await dialog.getByRole('button', { name: 'Import converted workbook' }).click();
    await expect(page.getByRole('tab', { name: 'Rule data', exact: true })).toBeVisible();
    expect(
      await page.evaluate(async () =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[5]]);
  });
}
