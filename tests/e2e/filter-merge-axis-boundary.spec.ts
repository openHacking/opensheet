import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('deleting filter columns and rejecting merged-axis edits preserve state and history', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Boundary', rows: 6, columns: 6 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    for (const seed of [71, 937, 9011]) {
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      const rows: (number | null)[][] = Array.from({ length: 5 }, (_, i) => [
        random(5),
        i + 10,
        random(3),
        i + 20,
        null,
        null,
      ]);
      rows.push([null, null, null, null, null, null]);
      const threshold = random(5);
      const firstFilter = {
        mode: 'all' as const,
        headerRows: 0,
        criteria: [
          { column: 0, operator: 'gt' as const, value: threshold },
          { column: 2, operator: 'neq' as const, value: 1 },
        ],
      };
      const secondFilter = {
        mode: 'all' as const,
        headerRows: 0,
        criteria: [{ column: 1, operator: 'neq' as const, value: 1 }],
      };
      const originalMerge = [{ startRow: 5, startColumn: 4, endRow: 6, endColumn: 6 }];
      const shiftedMerge = [{ startRow: 5, startColumn: 3, endRow: 6, endColumn: 5 }];
      const finalMerge = [{ startRow: 5, startColumn: 2, endRow: 6, endColumn: 4 }];
      const firstDeleted = rows.map((row) => row.slice(1));
      const secondDeleted = firstDeleted.map((row) => row.filter((_, index) => index !== 1));
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      await sheet.range('A1:F5').setValues(rows.slice(0, 5));
      await sheet.setFilter(firstFilter);
      await sheet.range('E6:F6').merge();
      const check = async (
        expected: (number | null)[][],
        filter: typeof firstFilter | typeof secondFilter | null,
        merges: typeof originalMerge,
        hidden: number[],
      ) => {
        const data = await sheet
          .range({ startRow: 0, endRow: 6, startColumn: 0, endColumn: expected[0].length })
          .getValues();
        const meta = book.sheetData(sheet.id);
        return (
          sheet.columnCount === expected[0].length &&
          JSON.stringify(data) === JSON.stringify(expected) &&
          JSON.stringify(meta.filter) === JSON.stringify(filter) &&
          JSON.stringify(meta.merges) === JSON.stringify(merges) &&
          JSON.stringify(book.filteredRows(sheet.id)) === JSON.stringify(hidden)
        );
      };
      const initialHidden = rows.flatMap((row, i) =>
        row[0] !== null && row[0] > threshold && row[2] !== null && row[2] !== 1 ? [] : [i],
      );
      const secondHidden = firstDeleted.flatMap((row, i) =>
        row[1] !== null && row[1] !== 1 ? [] : [i],
      );
      if (!(await check(rows, firstFilter, originalMerge, initialHidden)))
        return { seed, stage: 'initial' };
      await sheet.deleteColumns(0);
      if (!(await check(firstDeleted, secondFilter, shiftedMerge, secondHidden)))
        return { seed, stage: 'delete first criterion' };
      const before = {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        nextColumnId: book.sheetData(sheet.id).nextColumnId,
        columnOrder: [...book.sheetData(sheet.id).columnOrder],
      };
      for (const [kind, index] of [
        ['insert', 4],
        ['delete', 3],
      ] as const) {
        let code = '';
        try {
          if (kind === 'insert') await sheet.insertColumns(index);
          else await sheet.deleteColumns(index);
        } catch (error) {
          code = (error as { code?: string }).code ?? '';
        }
        if (
          code !== 'INVALID_RANGE' ||
          book.revision !== before.revision ||
          (await book.storageStats()).bytes !== before.bytes ||
          book.sheetData(sheet.id).nextColumnId !== before.nextColumnId ||
          JSON.stringify(book.sheetData(sheet.id).columnOrder) !==
            JSON.stringify(before.columnOrder) ||
          !(await check(firstDeleted, secondFilter, shiftedMerge, secondHidden))
        )
          return { seed, stage: `rejected ${kind}`, code };
      }
      if (!(await book.undo()) || !(await check(rows, firstFilter, originalMerge, initialHidden)))
        return { seed, stage: 'undo first deletion' };
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      if (!(await check(rows, firstFilter, originalMerge, initialHidden)))
        return { seed, stage: 'reopen before redo' };
      if (
        !(await book.redo()) ||
        !(await check(firstDeleted, secondFilter, shiftedMerge, secondHidden))
      )
        return { seed, stage: 'redo first deletion' };
      await sheet.deleteColumns(1);
      if (!(await check(secondDeleted, null, finalMerge, [])))
        return { seed, stage: 'delete last criterion' };
      if (
        !(await book.undo()) ||
        !(await check(firstDeleted, secondFilter, shiftedMerge, secondHidden))
      )
        return { seed, stage: 'undo last deletion' };
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      if (!(await book.redo()) || !(await check(secondDeleted, null, finalMerge, [])))
        return { seed, stage: 'redo last deletion after reopen' };
    }
    return null;
  }, file);
  expect(result).toBeNull();
});
