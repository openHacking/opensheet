import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('deleting filtered columns updates criteria and survives history and reopen', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Filter', rows: 5, columns: 8 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    type Criterion = { column: number; operator: 'gt' | 'lt'; value: number };
    type Model = {
      rows: (number | null)[][];
      criteria: Criterion[];
      mode: 'all' | 'any';
      mergeStart: number;
      ids: number[];
    };
    const app = window.opensheet;
    for (const mode of ['all', 'any'] as const) {
      for (const [first, second, combined] of [
        [0, 2, false],
        [1, 4, false],
        [2, 3, true],
      ] as const) {
        await app.load(JSON.parse(fileJSON));
        let book = app.getWorkbook(),
          sheet = book.getSheets()[0];
        let model: Model = {
          rows: Array.from({ length: 5 }, (_, row) =>
            Array.from({ length: 8 }, (_, column) =>
              row === 4 && column >= 6 ? null : row * 10 + column,
            ),
          ),
          criteria: [
            { column: first, operator: 'gt', value: 15 },
            { column: second, operator: 'lt', value: 35 },
          ],
          mode,
          mergeStart: 6,
          ids: [0, 1, 2, 3, 4, 5, 6, 7],
        };
        await sheet.range('A1:H5').setValues(model.rows);
        await sheet.setFilter({ mode, headerRows: 1, criteria: model.criteria });
        await sheet.range('G5:H5').merge();
        const clone = (m: Model): Model => ({
          ...m,
          rows: m.rows.map((r) => [...r]),
          criteria: m.criteria.map((c) => ({ ...c })),
          ids: [...m.ids],
        });
        const verify = async (stage: string) => {
          const meta = book.sheetData(sheet.id);
          const actual = await sheet
            .range({ startRow: 0, endRow: 5, startColumn: 0, endColumn: model.ids.length })
            .getValues();
          const hidden = model.criteria.length
            ? model.rows.flatMap((row, index) => {
                if (index === 0) return [];
                const matches = model.criteria.map(
                  (c) =>
                    row[c.column] !== null &&
                    (c.operator === 'gt' ? row[c.column]! > c.value : row[c.column]! < c.value),
                );
                return (mode === 'all' ? matches.every(Boolean) : matches.some(Boolean))
                  ? []
                  : [index];
              })
            : [];
          const filter = model.criteria.length
            ? { mode, headerRows: 1, criteria: model.criteria }
            : null;
          const merges = [
            {
              startRow: 4,
              startColumn: model.mergeStart,
              endRow: 5,
              endColumn: model.mergeStart + 2,
            },
          ];
          if (
            JSON.stringify(actual) !== JSON.stringify(model.rows) ||
            JSON.stringify(meta.filter) !== JSON.stringify(filter) ||
            JSON.stringify(book.filteredRows(sheet.id)) !== JSON.stringify(hidden) ||
            JSON.stringify(
              meta.merges.map((m) => [m.startRow, m.startColumn, m.endRow, m.endColumn]),
            ) !==
              JSON.stringify(
                merges.map((m) => [m.startRow, m.startColumn, m.endRow, m.endColumn]),
              ) ||
            JSON.stringify(meta.columnOrder) !== JSON.stringify(model.ids)
          )
            throw Error(
              JSON.stringify({
                mode,
                first,
                second,
                combined,
                stage,
                model,
                actual,
                filter: meta.filter,
                hidden: book.filteredRows(sheet.id),
                expectedHidden: hidden,
                merges: meta.merges,
                ids: meta.columnOrder,
              }),
            );
        };
        const remove = async (index: number, count: number) => {
          const before = clone(model);
          await sheet.deleteColumns(index, count);
          for (const row of model.rows) row.splice(index, count);
          model.ids.splice(index, count);
          model.mergeStart -= count;
          model.criteria = model.criteria
            .filter((c) => c.column < index || c.column >= index + count)
            .map((c) => ({
              ...c,
              column: c.column >= index + count ? c.column - count : c.column,
            }));
          const after = clone(model);
          await verify('delete');
          if (!(await book.undo())) throw Error('undo failed');
          model = before;
          await verify('undo');
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          await verify('reopen undo');
          if (!(await book.redo())) throw Error('redo failed');
          model = after;
          await verify('redo');
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          await verify('reopen redo');
        };
        await verify('initial');
        await remove(first, combined ? 2 : 1);
        if (!combined) await remove(model.criteria[0].column, 1);
        if (model.criteria.length) throw Error('all criteria should be removed');
      }
    }
    return 'ok';
  }, file);
  expect(result).toBe('ok');
});

test('deleting a filtered precedent updates surviving formula filters through history', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Formula filter', rows: 5, columns: 8 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    await app.load(JSON.parse(fileJSON));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('B2:B5').setValues([[10], [20], [30], [40]]);
    await sheet.range('F2:F5').setFormulas([['B2*2'], ['B3*2'], ['B4*2'], ['B5*2']]);
    await sheet.setFilter({
      mode: 'all',
      headerRows: 1,
      criteria: [
        { column: 1, operator: 'gt', value: 15 },
        { column: 5, operator: 'gt', value: 50 },
      ],
    });
    await sheet.range('G5:H5').merge();
    const observe = async (column: string) => ({
      values: await sheet.range(`${column}2:${column}5`).getValues(),
      formulas: await sheet.range(`${column}2:${column}5`).getFormulas(),
      hidden: book.filteredRows(sheet.id),
      filter: book.sheetData(sheet.id).filter,
      merges: book.sheetData(sheet.id).merges,
    });
    const before = await observe('F');
    await sheet.deleteColumns(1);
    const deleted = await observe('E');
    await book.undo();
    const undone = await observe('F');
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedUndo = await observe('F');
    await book.redo();
    const redone = await observe('E');
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedRedo = await observe('E');
    return { before, deleted, undone, reopenedUndo, redone, reopenedRedo };
  }, file);
  expect(result.before.values).toEqual([[20], [40], [60], [80]]);
  expect(result.before.hidden).toEqual([1, 2]);
  expect(result.undone).toEqual(result.before);
  expect(result.reopenedUndo).toEqual(result.before);
  expect(result.deleted.values).toEqual(Array.from({ length: 4 }, () => [{ error: '#REF!' }]));
  expect(result.deleted.formulas.flat().every((formula) => formula?.includes('#REF!'))).toBe(true);
  expect(result.deleted.hidden).toEqual([1, 2, 3, 4]);
  expect(result.deleted.filter).toEqual({
    mode: 'all',
    headerRows: 1,
    criteria: [{ column: 4, operator: 'gt', value: 50 }],
  });
  expect(result.deleted.merges).toEqual([{ startRow: 4, startColumn: 5, endRow: 5, endColumn: 7 }]);
  expect(result.redone).toEqual(result.deleted);
  expect(result.reopenedRedo).toEqual(result.deleted);
});
