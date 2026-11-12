import { test, expect } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('filter formulas use staged edits on another sheet through undo and reopen', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook();
      let [data, limits] = book.getSheets();
      await limits.range('A1').setValues([[5]]);
      await data.range('A1:A3').setValues([[1], [3], [7]]);
      await data
        .range('B1:B3')
        .setFormulas([['A1>Limits!$A$1'], ['A2>Limits!$A$1'], ['A3>Limits!$A$1']]);
      await data.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 1, operator: 'eq', value: true }],
      });
      const before = book.filteredRows(data.id);
      await limits.range('A1').setValues([[2]]);
      const edited = book.filteredRows(data.id);
      await book.undo();
      const undone = book.filteredRows(data.id);
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      data = book.getSheets()[0];
      return {
        before,
        edited,
        undone,
        reopened: book.filteredRows(data.id),
        values: await data.range('B1:B3').getValues(),
      };
    },
    JSON.stringify(
      createWorkbookFile({
        sheets: [
          { name: 'Data', rows: 3, columns: 2 },
          { name: 'Limits', rows: 3, columns: 2 },
        ],
      }),
    ),
  );
  expect(result).toEqual({
    before: [0, 1],
    edited: [0],
    undone: [0, 1],
    reopened: [0],
    values: [[false], [true], [true]],
  });
});

test('paged fill recalculates formula-based filter visibility through history and reopen', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      await sheet
        .range('A1:A130')
        .setValues(Array.from({ length: 130 }, (_, index) => [index + 1]));
      await sheet
        .range('B1:B130')
        .setFormulas(Array.from({ length: 130 }, (_, index) => [`A${index + 1}*2`]));
      await book.execute({
        type: 'core.sheet.filter',
        payload: {
          sheetId: sheet.id,
          filter: {
            mode: 'all',
            headerRows: 0,
            criteria: [{ column: 1, operator: 'gt', value: 2 }],
          },
        },
      });
      const before = book.filteredRows(sheet.id);
      await sheet.range('A1:A130').fillDown();
      const filled = book.filteredRows(sheet.id);
      const formulaValue = await sheet.range('B130').getValues();
      await book.undo();
      const undone = book.filteredRows(sheet.id);
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      return {
        before,
        filled,
        formulaValue,
        undone,
        reopened: book.filteredRows(sheet.id),
        reopenedValue: await sheet.range('B130').getValues(),
      };
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Filter', rows: 130, columns: 2 }] })),
  );
  expect(result.before).toEqual([0]);
  expect(result.filled).toEqual(Array.from({ length: 130 }, (_, index) => index));
  expect(result.formulaValue).toEqual([[2]]);
  expect(result.undone).toEqual([0]);
  expect(result.reopened).toEqual(result.filled);
  expect(result.reopenedValue).toEqual([[2]]);
});

test('custom filter applies typed conditions and survives edit/history/reopen/clear', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      const sheet = app.getWorkbook().getSheets()[0];
      await sheet.range('A1:B5').setValues([
        ['Group', 'Score'],
        ['a', 1],
        ['b', 4],
        ['A', 3],
        ['a', '5'],
      ]);
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 5,
        startColumn: 0,
        endColumn: 2,
      });
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Filter', rows: 5 }] })),
  );
  await page.getByLabel('Sheet actions').selectOption('filter-custom');
  await page.getByLabel('Condition 1 operator').selectOption('eq');
  await page.getByLabel('Condition 1 value', { exact: true }).fill('a');
  await page.getByRole('button', { name: 'Add condition' }).click();
  await page.getByLabel('Condition 2 column').selectOption('1');
  await page.getByLabel('Condition 2 operator').selectOption('gt');
  await page.getByLabel('Condition 2 value', { exact: true }).fill('2');
  await page.getByRole('button', { name: 'Apply filter' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const initial = book.filteredRows(sheet.id);
    await sheet.range('B2').setValues([[3]]);
    const edited = book.filteredRows(sheet.id);
    await book.undo();
    const undone = book.filteredRows(sheet.id);
    await book.redo();
    await app.open(book.id);
    return { initial, edited, undone, reopened: app.getWorkbook().filteredRows(sheet.id) };
  });
  expect(result).toEqual({
    initial: [1, 2, 4],
    edited: [2, 4],
    undone: [1, 2, 4],
    reopened: [2, 4],
  });
  await page.getByLabel('Sheet actions').selectOption('filter-custom');
  await expect(page.getByLabel('Condition 2 operator')).toHaveValue('gt');
  await expect(page.getByLabel('Condition 2 value', { exact: true })).toHaveValue('2');
  await page.getByRole('button', { name: 'Clear all filters' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      return book.filteredRows(book.getSheets()[0].id);
    }),
  ).toEqual([]);
});

test('filter dialog rejects stale changes and read-only mutations', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.getByLabel('Sheet actions').selectOption('filter-custom');
  const revision = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['changed']]);
    return book.revision;
  });
  await page.getByRole('button', { name: 'Apply filter' }).click();
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Close this dialog and open the filter again.',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.keyboard.press('Escape');
  await page.getByLabel('Read only', { exact: true }).check();
  await page.getByLabel('Sheet actions').selectOption('filter-custom');
  await expect(page.getByRole('dialog')).not.toBeVisible();
});
