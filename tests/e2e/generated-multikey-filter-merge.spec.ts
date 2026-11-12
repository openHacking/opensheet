import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('seeded multi-key sort, filter and merge interactions preserve rows and history', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Model', rows: 9, columns: 3 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    type Row = [number, number, string];
    for (const [caseIndex, seed] of [83, 727, 4099, 65537].entries()) {
      let randomState = seed;
      const random = (n: number) => {
        randomState = (Math.imul(randomState, 1664525) + 1013904223) >>> 0;
        return Math.floor((randomState / 0x100000000) * n);
      };
      const original: Row[] = Array.from({ length: 8 }, (_, i) => [
        random(3),
        random(3),
        `id-${i}`,
      ]);
      original[0] = [1, 2, 'id-0'];
      original[1] = [1, 2, 'id-1'];
      const threshold = random(3);
      const firstDirection: 'asc' | 'desc' = caseIndex & 1 ? 'desc' : 'asc';
      const secondDirection: 'asc' | 'desc' = caseIndex & 2 ? 'desc' : 'asc';
      const sorted = original
        .map((row, index) => ({ row, index }))
        .sort(
          (a, b) =>
            (firstDirection === 'asc' ? a.row[0] - b.row[0] : b.row[0] - a.row[0]) ||
            (secondDirection === 'asc' ? a.row[1] - b.row[1] : b.row[1] - a.row[1]) ||
            a.index - b.index,
        )
        .map(({ row }) => row);
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      const keys = [
        { column: 0, direction: firstDirection },
        { column: 1, direction: secondDirection },
      ];
      const check = async (rows: Row[], merged: boolean) => {
        const actual = await sheet.range('A1:C8').getValues();
        const expectedHidden = rows.flatMap((row, i) =>
          row[0] >= threshold && row[1] !== 1 ? [] : [i],
        );
        expectedHidden.push(8);
        const hidden = book.filteredRows(sheet.id);
        const merges = book.sheetData(sheet.id).merges;
        const expectedMerges = merged
          ? [{ startRow: 8, startColumn: 0, endRow: 9, endColumn: 2 }]
          : [];
        return (
          JSON.stringify(actual) === JSON.stringify(rows) &&
          JSON.stringify(hidden) === JSON.stringify(expectedHidden) &&
          JSON.stringify(merges) === JSON.stringify(expectedMerges)
        );
      };
      await sheet.range('A1:C8').setValues(original);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [
          { column: 0, operator: 'gte', value: threshold },
          { column: 1, operator: 'neq', value: 1 },
        ],
      });
      await sheet.range('A9:B9').merge();
      const revision = book.revision;
      let error = '';
      try {
        await sheet.sort('A1:C9', keys);
      } catch (reason) {
        error = (reason as { code?: string }).code ?? '';
      }
      if (error !== 'INVALID_RANGE' || revision !== book.revision || !(await check(original, true)))
        return { seed, stage: 'merge rejection', error };
      await sheet.range('A9:B9').unmerge();
      await sheet.sort('A1:C8', keys);
      if (!(await check(sorted, false)))
        return {
          seed,
          stage: 'sort',
          original,
          sorted,
          actual: await sheet.range('A1:C8').getValues(),
          hidden: book.filteredRows(sheet.id),
        };
      if (!(await book.undo()) || !(await check(original, false)))
        return { seed, stage: 'undo sort' };
      if (!(await book.undo()) || !(await check(original, true)))
        return { seed, stage: 'undo unmerge' };
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      if (!(await check(original, true))) return { seed, stage: 'reopen' };
      if (!(await book.redo()) || !(await book.redo()) || !(await check(sorted, false)))
        return { seed, stage: 'redo after reopen' };
    }
    return null;
  }, file);
  expect(result).toBeNull();
});
