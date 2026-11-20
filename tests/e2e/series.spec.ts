import { workbookTool } from './workbook-tools.js';
import { test, expect } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('paged fill validates targets and composes atomically with transaction commands', async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([[1], [2], [3]]);
    const revision = book.revision;
    const invalid = await book
      .execute({
        type: 'core.cells.fillDown',
        payload: {
          sheetId: 'missing-sheet',
          range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
        },
      })
      .then(
        () => 'success',
        (error) => error.code,
      );
    const invalidRevision = book.revision;
    await book.transaction({}, async (tx) => {
      const target = tx.getSheetById(sheet.id)!;
      await target.range('A1').setValues([[9]]);
      await target.range('A1:A3').fillDown();
      await target.range('A3').setValues([[12]]);
    });
    const filled = await sheet.range('A1:A3').getValues();
    await book.undo();
    const undone = await sheet.range('A1:A3').getValues();
    await sheet.range('A2:B2').merge();
    const mergedRevision = book.revision;
    const merged = await sheet
      .range('A1:A3')
      .fillDown()
      .then(
        () => 'success',
        (error) => error.code,
      );
    return {
      invalid,
      revision,
      invalidRevision,
      filled,
      undone,
      merged,
      mergedRevision,
      finalRevision: book.revision,
      final: await sheet.range('A1:A3').getValues(),
    };
  });
  expect(result.invalid).toBe('INVALID_ARGUMENT');
  expect(result.invalidRevision).toBe(result.revision);
  expect(result.filled).toEqual([[9], [9], [12]]);
  expect(result.undone).toEqual([[1], [2], [3]]);
  expect(result.merged).toBe('INVALID_RANGE');
  expect(result.finalRevision).toBe(result.mergedRevision);
  expect(result.final).toEqual(result.undone);
});

test('paged fill preserves source formatting and annotations across pages without touching neighbors', async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.insertRows(100, 30);
    await sheet.range('A1:D130').clear({ all: true });
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('B1').setStyle({ bold: true, color: '#123456' });
    await sheet.range('B1').setNote('source');
    await sheet.range('B1').setLink({ target: 'https://example.com/' });
    await sheet.range('C1').setFormulas([['$A1+A$1']]);
    await sheet.range('B70').setValues([['old']]);
    await sheet.range('B70').setNote('old note');
    await sheet.range('D70').setValues([['neighbor']]);
    const read = async () => {
      const cell = await book.getCell(sheet.id, 69, 1);
      const formula = await book.getCell(sheet.id, 69, 2);
      return {
        values: await sheet.range('A70:D70').getValues(),
        note: (await sheet.range('B70').getNotes())[0][0],
        link: (await sheet.range('B70').getLinks())[0][0],
        style: book.getStyle(cell?.styleId),
        input: formula?.input,
      };
    };
    await sheet.range('A1:C130').fillDown();
    const filled = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { filled, undone, reopened: await read() };
  });
  expect(result.filled).toEqual({
    values: [[7, null, 14, 'neighbor']],
    note: 'source',
    link: { target: 'https://example.com/' },
    style: { bold: true, color: '#123456' },
    input: { type: 'formula', expression: '($A70+A$1)' },
  });
  expect(result.undone.values).toEqual([[null, 'old', null, 'neighbor']]);
  expect(result.undone.note).toBe('old note');
  expect(result.undone.link).toBeNull();
  expect(result.reopened).toEqual(result.filled);
});

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:B4').clear();
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 });
  });
});

test('fill down uses the first row and survives undo, redo and reopen', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('C1:C3').clear();
    await sheet.range('A1:B3').setValues([
      [2, null],
      [9, 'old'],
      [8, 'old2'],
    ]);
    await sheet.range('C1').setFormulas([['A1*2']]);
    await sheet.range('A1:C3').fillDown();
    const filled = await sheet.range('A1:C3').getValues();
    await book.undo();
    const undone = await sheet.range('A1:C3').getValues();
    await book.redo();
    await app.open(book.id);
    return {
      filled,
      undone,
      reopened: await app.getWorkbook().getSheets()[0].range('A1:C3').getValues(),
    };
  });
  expect(result.filled).toEqual([
    [2, null, 4],
    [2, null, 4],
    [2, null, 4],
  ]);
  expect(result.undone).toEqual([
    [2, null, 4],
    [9, 'old', null],
    [8, 'old2', null],
  ]);
  expect(result.reopened).toEqual(result.filled);
});

test('fill down follows reordered physical row and column ids after insertion', async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.insertRows(1);
    await sheet.insertColumns(1);
    await sheet.range('A1:C4').clear({ all: true });
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('C1').setFormulas([['A1*2']]);
    await sheet.range('B2:B4').setValues([['old'], ['old2'], ['old3']]);
    await sheet.range('A1:C4').fillDown();
    const filled = await sheet.range('A1:C4').getValues();
    await book.undo();
    const undone = await sheet.range('A1:C4').getValues();
    await book.redo();
    await app.open(book.id);
    return {
      filled,
      undone,
      reopened: await app.getWorkbook().getSheets()[0].range('A1:C4').getValues(),
    };
  });
  expect(result.filled).toEqual(Array.from({ length: 4 }, () => [7, null, 14]));
  expect(result.undone).toEqual([
    [7, null, 14],
    [null, 'old', null],
    [null, 'old2', null],
    [null, 'old3', null],
  ]);
  expect(result.reopened).toEqual(result.filled);
});

test('fills decimal series and restores it through history and reopen', async ({ page }) => {
  await workbookTool(page, 'fill-series');
  await expect(page.getByLabel('Start value')).toBeEditable();
  await page.getByLabel('Start value').fill('0.1');
  await page.getByLabel('Step', { exact: true }).fill('0.1');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      range = book.getSheets()[0].range('A1:B4');
    const filled = await range.getValues();
    await book.undo();
    const undone = await range.getValues();
    await book.redo();
    await app.open(book.id);
    return {
      filled,
      undone,
      restored: await app.getWorkbook().getSheets()[0].range('A1:B4').getValues(),
    };
  });
  expect(result.filled).toEqual([
    [0.1, 0.1],
    [0.2, 0.2],
    [0.3, 0.3],
    [0.4, 0.4],
  ]);
  expect(result.undone).toEqual(Array.from({ length: 4 }, () => [null, null]));
  expect(result.restored).toEqual(result.filled);
});

test('reports overflow and stale revisions without overwriting cells', async ({ page }) => {
  await workbookTool(page, 'fill-series');
  await expect(page.getByLabel('Start value')).toBeEditable();
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByLabel('Start value').fill('1e308');
  await page.getByLabel('Step', { exact: true }).fill('1e308');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('finite numeric range');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.getByLabel('Start value').fill('1');
  await page.getByLabel('Step', { exact: true }).fill('1');
  const changed = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[99]]);
    return book.revision;
  });
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Close this dialog and select the range again.',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(changed);
});

test('fills across columns and disables filling in read-only mode', async ({ page }) => {
  await workbookTool(page, 'fill-series');
  await expect(page.getByLabel('Start value')).toBeEditable();
  await page.getByLabel('Start value').fill('3');
  await page.getByLabel('Step', { exact: true }).fill('-2');
  await page.getByLabel('Direction', { exact: true }).selectOption('column');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:B4').getValues(),
    ),
  ).toEqual(Array.from({ length: 4 }, () => [3, 1]));
  await page.getByLabel('Read only', { exact: true }).check();
  await workbookTool(page, 'fill-series', true);
  await expect(page.getByRole('dialog')).not.toBeVisible();
});

test('fills dates atomically with formatting, undo, redo and reopen', async ({ page }) => {
  await page.evaluate(async () => {
    const range = window.opensheet.getWorkbook().getSheets()[0].range('A1:B4');
    await range.clear({ all: true });
    await range.setStyle({ bold: true, numberFormat: '0.00' });
    await range.setNote('keep');
  });
  await workbookTool(page, 'fill-series');
  await page.getByLabel('Series type').selectOption('day');
  await page.getByLabel('Start date').fill('1900-02-28');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const read = async () => ({
      values: await sheet.range('A1:B4').getValues(),
      display: await sheet.range('A1:B4').getDisplayValues(),
      note: (await sheet.range('A1').getNotes())[0][0],
      style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
    });
    const filled = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { filled, undone, reopened: await read() };
  });
  expect(result.filled.values).toEqual([
    [59, 59],
    [60, 60],
    [61, 61],
    [62, 62],
  ]);
  expect(result.filled.display).toEqual([
    ['1900-02-28', '1900-02-28'],
    ['1900-02-29', '1900-02-29'],
    ['1900-03-01', '1900-03-01'],
    ['1900-03-02', '1900-03-02'],
  ]);
  expect(result.filled.style).toEqual({ bold: true, numberFormat: 'yyyy-mm-dd' });
  expect(result.filled.note).toBe('keep');
  expect(result.undone.values).toEqual(Array.from({ length: 4 }, () => [null, null]));
  expect(result.undone.style).toEqual({ bold: true, numberFormat: '0.00' });
  expect(result.undone.note).toBe('keep');
  expect(result.reopened).toEqual(result.filled);
});

test('rejects date series outside the date system without changing the workbook', async ({
  page,
}) => {
  await workbookTool(page, 'fill-series');
  await page.getByLabel('Series type').selectOption('day');
  await page.getByLabel('Start date').fill('9999-12-31');
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Dates must stay within');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.getByLabel('Start date').fill('1900-01-01');
  await page.getByLabel('Step', { exact: true }).fill('-1');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Dates must stay within');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
});

test('fills descending dates across columns in the 1904 system', async ({ page }) => {
  const file = {
    ...createWorkbookFile({ sheets: [{ name: 'Dates', rows: 100 }] }),
    dateSystem: '1904' as const,
  };
  await page.evaluate(async (input) => {
    const app = window.opensheet;
    await app.load(JSON.parse(input));
    const sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 3 });
  }, JSON.stringify(file));
  await workbookTool(page, 'fill-series');
  await page.getByLabel('Series type').selectOption('day');
  await page.getByLabel('Start date').fill('1904-01-03');
  await page.getByLabel('Step', { exact: true }).fill('-1');
  await page.getByLabel('Direction').selectOption('column');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const range = book.getSheets()[0].range('A1:C2');
    const values = await range.getValues();
    const display = await range.getDisplayValues();
    await app.open(book.id);
    return {
      values,
      display,
      reopened: await app.getWorkbook().getSheets()[0].range('A1:C2').getDisplayValues(),
    };
  });
  expect(result.values).toEqual([
    [2, 1, 0],
    [2, 1, 0],
  ]);
  expect(result.display).toEqual([
    ['1904-01-03', '1904-01-02', '1904-01-01'],
    ['1904-01-03', '1904-01-02', '1904-01-01'],
  ]);
  expect(result.reopened).toEqual(result.display);
});

test('date series rejects stale revisions and merged cells without partial changes', async ({
  page,
}) => {
  await workbookTool(page, 'fill-series');
  await page.getByLabel('Series type').selectOption('day');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[42]]);
  });
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Workbook changed');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.getByRole('dialog').getByRole('button', { name: 'Close' }).click();
  await page.evaluate(async () => {
    await window.opensheet.getWorkbook().getSheets()[0].range('A2:B2').merge();
  });
  await workbookTool(page, 'fill-series');
  await page.getByLabel('Series type').selectOption('day');
  const mergedRevision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('alert')).toContainText('Unmerge');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(mergedRevision);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[42]]);
});

test('fills calendar months from the original day and persists undo and redo', async ({ page }) => {
  await workbookTool(page, 'fill-series');
  await page.getByLabel('Series type').selectOption('month');
  await page.getByLabel('Start date').fill('2024-01-31');
  await page.getByRole('button', { name: 'Fill series', exact: true }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let range = book.getSheets()[0].range('A1:B4');
    const filled = await range.getDisplayValues();
    await book.undo();
    const undone = await range.getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    range = book.getSheets()[0].range('A1:B4');
    return { filled, undone, reopened: await range.getDisplayValues() };
  });
  expect(result.filled).toEqual([
    ['2024-01-31', '2024-01-31'],
    ['2024-02-29', '2024-02-29'],
    ['2024-03-31', '2024-03-31'],
    ['2024-04-30', '2024-04-30'],
  ]);
  expect(result.undone).toEqual(Array.from({ length: 4 }, () => [null, null]));
  expect(result.reopened).toEqual(result.filled);
});

test('supports year series through the public API and rejects overflow atomically', async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const range = book.getSheets()[0].range('A1:A4');
    await range.fillSeries({ start: 43890, step: 1, calendar: 'year' });
    const values = await range.getValues();
    const revision = book.revision;
    const failure = await range.fillSeries({ start: 2958465, step: 1, calendar: 'year' }).then(
      () => 'success',
      (error) => error.code,
    );
    const after = await range.getValues();
    return { values, revision, finalRevision: book.revision, failure, after };
  });
  expect(result.values).toEqual([[43890], [44255], [44620], [44985]]);
  expect(result.failure).toBe('INVALID_ARGUMENT');
  expect(result.finalRevision).toBe(result.revision);
  expect(result.after).toEqual(result.values);
});
