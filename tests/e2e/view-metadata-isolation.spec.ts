import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('used range and filtered rows cannot be mutated through a returned view', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async (file) => {
      const app = window.opensheet;
      await app.load(JSON.parse(file));
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      await sheet.range('A1:A5').setValues([[2], [1], [3], [0], [4]]);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 0, operator: 'gte', value: 2 }],
      });
      const revision = book.revision;
      const inspect = () => {
        const used = book.usedRange(sheet.id)!;
        const rows = book.filteredRows(sheet.id);
        const rangeWrite = Reflect.set(used, 'endRow', 1);
        const rowWrite = Reflect.set(rows, '0', 4);
        return {
          rangeWrite,
          rowWrite,
          used: book.usedRange(sheet.id),
          rows: book.filteredRows(sheet.id),
          sheetUsed: sheet.getUsedRange().bounds,
        };
      };
      const original = inspect();
      await sheet.setRowsHidden(0, 1, false);
      const afterNoop = inspect();
      const afterNoopRevision = book.revision;
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      const reopened = inspect();
      return {
        original,
        afterNoop,
        reopened,
        revision,
        afterNoopRevision,
        reopenedRevision: book.revision,
        values: await sheet.range('A1:A5').getValues(),
      };
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Data', rows: 5, columns: 2 }] })),
  );
  const expected = {
    rangeWrite: false,
    rowWrite: false,
    used: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 1 },
    rows: [1, 3],
    sheetUsed: { startRow: 0, startColumn: 0, endRow: 5, endColumn: 1 },
  };
  expect(result.original).toEqual(expected);
  expect(result.afterNoop).toEqual(expected);
  expect(result.reopened).toEqual(expected);
  expect(result.afterNoopRevision).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.values).toEqual([[2], [1], [3], [0], [4]]);
});
