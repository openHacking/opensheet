import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

test('reports every independently authored defined name before import and after reopen', async ({
  page,
}) => {
  const root = new URL('../../benchmarks/', import.meta.url);
  const bytes = readFileSync(new URL('xlsx-defined-names-openpyxl-2026-11-12.xlsx', root));
  const fixture = JSON.parse(
    readFileSync(new URL('xlsx-defined-names-openpyxl-2026-11-12.xlsx.json', root), 'utf8'),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  expect(fixture.globalNames).toEqual(['GlobalValue', 'OtherValue']);
  expect(fixture.localNames).toEqual(['LocalValue']);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => ({
    id: window.opensheet.getWorkbook().id,
    revision: window.opensheet.getWorkbook().revision,
  }));
  await page.locator('#file').setInputFiles({
    name: 'defined-names.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(report).toContainText('Defined names are not preserved (3 occurrences)');
  expect(
    await page.evaluate(() => ({
      id: window.opensheet.getWorkbook().id,
      revision: window.opensheet.getWorkbook().revision,
    })),
  ).toEqual(before);
  await report.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Named data', exact: true })).toBeVisible();
  const values = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const initial = await sheet.range('A1:A3').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { initial, reopened: await sheet.range('A1:A3').getValues() };
  });
  expect(values).toEqual({ initial: [[11], [22], [33]], reopened: [[11], [22], [33]] });
  await page.locator('#export-xlsx').click();
  await expect(report).toContainText('Defined names are not preserved (3 occurrences)');
});

test('locates formulas that use imported defined names before accepting the conversion', async ({
  page,
}) => {
  const root = new URL('../../benchmarks/', import.meta.url);
  const bytes = readFileSync(new URL('xlsx-defined-name-formulas-openpyxl-2026-11-12.xlsx', root));
  const fixture = JSON.parse(
    readFileSync(new URL('xlsx-defined-name-formulas-openpyxl-2026-11-12.xlsx.json', root), 'utf8'),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  expect(fixture.formulas).toEqual([
    '=GlobalValue*2',
    '=LocalValue+OtherValue',
    '="GlobalValue"',
    '=SUM(A1:A3)',
  ]);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'name-formulas.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(report).toContainText(
    'Named data: B1: Formula references a defined name that is not preserved (2 occurrences; first location shown)',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await report.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Named data', exact: true })).toBeVisible();
  const observed = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const read = async () => ({
      values: await sheet.range('B1:B4').getValues(),
      formulas: await sheet.range('B1:B4').getFormulas(),
    });
    const initial = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { initial, reopened: await read() };
  });
  expect(observed.initial.values).toEqual([
    [{ error: '#NAME?' }],
    [{ error: '#NAME?' }],
    ['GlobalValue'],
    [66],
  ]);
  expect(observed.initial.formulas).toEqual([
    ['GlobalValue*2'],
    ['LocalValue+OtherValue'],
    ['"GlobalValue"'],
    ['SUM(A1:A3)'],
  ]);
  expect(observed.reopened).toEqual(observed.initial);
  await page.locator('#export-xlsx').click();
  await expect(report).toContainText(
    'Formula references a defined name that is not preserved (2 occurrences',
  );
});

test('reports a sheet-qualified local name without confusing it with a cell or external reference', async ({
  page,
}) => {
  const root = new URL('../../benchmarks/', import.meta.url);
  const bytes = readFileSync(new URL('xlsx-qualified-names-openpyxl-2026-11-12.xlsx', root));
  const fixture = JSON.parse(
    readFileSync(new URL('xlsx-qualified-names-openpyxl-2026-11-12.xlsx.json', root), 'utf8'),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'qualified.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  const named = report.locator('li', {
    hasText: 'Formula references a defined name that is not preserved',
  });
  await expect(named).toHaveCount(1);
  await expect(named).toContainText('Other: A1:');
  await expect(named).not.toContainText('occurrences');
  await expect(report).toContainText(
    'External workbook and structured table references are not evaluated',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await report.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheetByName('Other')!;
    const read = async () => ({
      formulas: await sheet.range('A1:A3').getFormulas(),
      values: await sheet.range('A1:A3').getValues(),
    });
    const initial = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheetByName('Other')!;
    return { initial, reopened: await read() };
  });
  expect(result.initial.formulas).toEqual([
    ["'Named data'!LocalValue+1"],
    ["'Named data'!$A$1"],
    ["'[outside.xlsx]Named data'!LocalValue"],
  ]);
  expect(result.initial.values[1]).toEqual([11]);
  expect(result.initial.values[0][0]).toEqual(
    expect.objectContaining({ error: expect.any(String) }),
  );
  expect(result.reopened).toEqual(result.initial);
});

test('reports a qualified workbook name before import and keeps the report after reopen', async ({
  page,
}) => {
  const root = new URL('../../benchmarks/', import.meta.url);
  const bytes = readFileSync(new URL('xlsx-qualified-global-openpyxl-2026-11-12.xlsx', root));
  const fixture = JSON.parse(
    readFileSync(new URL('xlsx-qualified-global-openpyxl-2026-11-12.xlsx.json', root), 'utf8'),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'qualified-global.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  const named = report.locator('li', {
    hasText: 'Formula references a defined name that is not preserved',
  });
  await expect(named).toHaveCount(1);
  await expect(named).toContainText('Other: A1:');
  await expect(named).not.toContainText('occurrences');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await report.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const initial = await book.getSheetByName('Other')!.range('A1:A4').getFormulas();
    await app.open(book.id);
    book = app.getWorkbook();
    return {
      initial,
      reopened: await book.getSheetByName('Other')!.range('A1:A4').getFormulas(),
      file: await book.toJSON(),
    };
  });
  expect(result.initial).toEqual([
    ["'Named data'!GlobalValue+1"],
    ["'Named data'!$A$1"],
    ['"GlobalValue"'],
    ["'[outside.xlsx]Named data'!GlobalValue"],
  ]);
  expect(result.reopened).toEqual(result.initial);
  expect(JSON.stringify(result.file)).toContain('DEFINED_NAME_REFERENCE_DROPPED');
});

test('reports Unicode defined-name references before import and retains formulas after reopen', async ({
  page,
}) => {
  const root = new URL('../../benchmarks/', import.meta.url);
  const bytes = readFileSync(new URL('xlsx-unicode-names-openpyxl-2026-11-12.xlsx', root));
  const fixture = JSON.parse(
    readFileSync(new URL('xlsx-unicode-names-openpyxl-2026-11-12.xlsx.json', root), 'utf8'),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'unicode-names.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  const named = report.locator('li', {
    hasText: 'Formula references a defined name that is not preserved',
  });
  await expect(named).toHaveCount(2);
  await expect(named.nth(0)).toContainText('数据: B1:');
  await expect(named.nth(0)).toContainText('2 occurrences');
  await expect(named.nth(1)).toContainText('Other: A1:');
  await expect(named.nth(1)).toContainText('2 occurrences');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await report.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const read = async () => ({
      data: await book.getSheetByName('数据')!.range('B1:B3').getFormulas(),
      other: await book.getSheetByName('Other')!.range('A1:A4').getFormulas(),
    });
    const initial = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    return { initial, reopened: await read(), file: await book.toJSON() };
  });
  expect(result.initial.data).toEqual([['预算+1'], ['总计+1'], ['"预算"']]);
  expect(result.initial.other).toEqual([
    ["'数据'!总计+1"],
    ["'数据'!预算+1"],
    ["'数据'!$A$1"],
    ["'[external.xlsx]数据'!预算"],
  ]);
  expect(result.reopened).toEqual(result.initial);
  expect(JSON.stringify(result.file)).toContain('DEFINED_NAME_REFERENCE_DROPPED');
});

test('reports backslash-prefixed names before import and retains the report', async ({ page }) => {
  const root = new URL('../../benchmarks/', import.meta.url);
  const bytes = readFileSync(new URL('xlsx-backslash-names-openpyxl-2026-11-12.xlsx', root));
  const fixture = JSON.parse(
    readFileSync(new URL('xlsx-backslash-names-openpyxl-2026-11-12.xlsx.json', root), 'utf8'),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => window.opensheet.getWorkbook().id);
  await page.locator('#file').setInputFiles({
    name: 'backslash-names.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: bytes,
  });
  const report = page.getByRole('dialog', { name: 'Compatibility report' });
  const named = report.locator('li', {
    hasText: 'Formula references a defined name that is not preserved',
  });
  await expect(named).toHaveCount(2);
  await expect(named.nth(0)).toContainText('Data: B1:');
  await expect(named.nth(0)).toContainText('2 occurrences');
  await expect(named.nth(1)).toContainText('Other: A1:');
  await expect(named.nth(1)).toContainText('2 occurrences');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(before);
  await report.getByRole('button', { name: 'Import converted workbook' }).click();
  await expect(page.getByRole('tab', { name: 'Other', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const read = async () => ({
      data: await book.getSheetByName('Data')!.range('B1:B3').getFormulas(),
      other: await book.getSheetByName('Other')!.range('A1:A4').getFormulas(),
    });
    const initial = await read();
    await app.open(book.id);
    book = app.getWorkbook();
    return { initial, reopened: await read(), file: await book.toJSON() };
  });
  expect(result.initial.data).toEqual([['\\Budget+1'], ['\\本地+1'], ['"\\Budget"']]);
  expect(result.initial.other).toEqual([
    ["'Data'!\\Budget+1"],
    ["'Data'!\\本地+1"],
    ["'Data'!$A$1"],
    ["'[external.xlsx]Data'!\\Budget"],
  ]);
  expect(result.reopened).toEqual(result.initial);
  expect(JSON.stringify(result.file)).toContain('DEFINED_NAME_REFERENCE_DROPPED');
});
