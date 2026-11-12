import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('merge rejects sorting atomically while filtering and history survive reopen', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Interactions', rows: 4, columns: 2 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    const fixtures: { values: (number | null)[]; threshold: number; direction: 'asc' | 'desc' }[] =
      [
        { values: [3, 1, 2, 4], threshold: 2, direction: 'asc' },
        { values: [2, null, 2, 1], threshold: 1, direction: 'desc' },
        { values: [null, 4, 0, 4], threshold: 3, direction: 'asc' },
      ];
    for (const fixture of fixtures) {
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      const rows: [number | null, string | null][] = fixture.values.map((value, index) => [
        value,
        index ? `row-${index}` : null,
      ]);
      const check = async (expected: [number | null, string | null][], merged: boolean) => {
        const actual = await sheet.range('A1:B4').getValues();
        const values = expected.map((row) => row[0]);
        const hidden = book.filteredRows(sheet.id);
        const expectedHidden = values.flatMap((value, index) =>
          value === null || value <= fixture.threshold ? [index] : [],
        );
        return (
          JSON.stringify(actual) === JSON.stringify(expected) &&
          JSON.stringify(hidden) === JSON.stringify(expectedHidden) &&
          JSON.stringify(book.sheetData(sheet.id).merges) ===
            JSON.stringify(merged ? [{ startRow: 0, startColumn: 0, endRow: 1, endColumn: 2 }] : [])
        );
      };
      await sheet.range('A1:B4').setValues(rows);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 0, operator: 'gt', value: fixture.threshold }],
      });
      await sheet.range('A1:B1').merge();
      const revision = book.revision;
      let error = '';
      try {
        await sheet.sort('A1:B4', 0, fixture.direction);
      } catch (e) {
        error = (e as { code?: string }).code ?? '';
      }
      if (error !== 'INVALID_RANGE' || book.revision !== revision || !(await check(rows, true)))
        return { fixture, stage: 'rejected sort', error, revision, actualRevision: book.revision };
      await sheet.range('A1:B1').unmerge();
      await sheet.sort('A1:B4', 0, fixture.direction);
      const sorted = rows
        .map((row, index) => ({ value: row[0], row, index }))
        .sort((a, b) => {
          if (a.value === null && b.value === null) return a.index - b.index;
          if (a.value === null) return 1;
          if (b.value === null) return -1;
          return (
            (fixture.direction === 'asc' ? a.value - b.value : b.value - a.value) ||
            a.index - b.index
          );
        })
        .map((item) => item.row);
      if (!(await check(sorted, false)))
        return { fixture, stage: 'sorted', sorted, hidden: book.filteredRows(sheet.id) };
      await book.undo();
      if (!(await check(rows, false))) return { fixture, stage: 'undo sort' };
      await book.undo();
      if (!(await check(rows, true))) return { fixture, stage: 'undo unmerge' };
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      if (!(await check(rows, true))) return { fixture, stage: 'reopen' };
      await book.redo();
      await book.redo();
      if (!(await check(sorted, false)))
        return { fixture, stage: 'redo after reopen', hidden: book.filteredRows(sheet.id) };
    }
    return null;
  }, file);
  expect(result).toBeNull();
});
