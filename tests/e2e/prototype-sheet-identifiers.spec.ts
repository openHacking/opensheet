import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

for (const id of ['toString', 'valueOf', 'hasOwnProperty', '__defineGetter__']) {
  test(`sheet ID ${id} has ordinary used ranges and filter behavior`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 5, columns: 2 }] });
    file.sheets[0].id = id;
    const result = await page.evaluate(async (inputJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(inputJSON));
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      const empty = { used: book.usedRange(sheet.id), filtered: book.filteredRows(sheet.id) };
      const unknown = ['__proto__', 'constructor', 'prototype'].map((key) => ({
        used: book.usedRange(key),
        filtered: book.filteredRows(key),
      }));
      await sheet.range('A1:A5').setValues([[2], [1], [3], [0], [4]]);
      const navigation = await book.navigateData(sheet.id, 0, 0, 'row', 1);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 0, operator: 'gte', value: 2 }],
      });
      const filtered = book.filteredRows(sheet.id);
      const used = book.usedRange(sheet.id);
      const filteredNavigation = await book.navigateData(sheet.id, 0, 0, 'row', 1);
      await book.undo();
      const undone = book.filteredRows(sheet.id);
      await book.redo();
      const redone = book.filteredRows(sheet.id);
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      return {
        empty,
        unknown,
        navigation,
        filteredNavigation,
        undone,
        redone,
        used,
        filtered,
        reopened: {
          id: sheet.id,
          used: book.usedRange(sheet.id),
          filtered: book.filteredRows(sheet.id),
          values: await sheet.range('A1:A5').getValues(),
        },
      };
    }, JSON.stringify(file));
    const used = { startRow: 0, startColumn: 0, endRow: 5, endColumn: 1 };
    expect(result.empty).toEqual({ used: { ...used, endRow: 1 }, filtered: [] });
    expect(result.unknown).toEqual(
      Array.from({ length: 3 }, () => ({ used: undefined, filtered: [] })),
    );
    expect(result.navigation).toMatchObject({ row: 4, column: 0 });
    expect(result.filteredNavigation).toMatchObject({ row: 4, column: 0 });
    expect(result.used).toEqual(used);
    expect(result.filtered).toEqual([1, 3]);
    expect(result.undone).toEqual([]);
    expect(result.redone).toEqual([1, 3]);
    expect(result.reopened).toEqual({
      id,
      used,
      filtered: [1, 3],
      values: [[2], [1], [3], [0], [4]],
    });
  });
}
