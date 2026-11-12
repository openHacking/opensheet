import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('seeded edits, axis changes, sort, filter and merge agree with a model through history and reopen', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Combined', rows: 6, columns: 3 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    type Row = (number | null)[];
    type Model = { rows: Row[]; threshold: number | null; merged: boolean; key: number };
    const app = window.opensheet;
    for (const seed of [61, 1907, 9119]) {
      await app.load(JSON.parse(fileJSON));
      let model: Model = {
        rows: Array.from({ length: 6 }, () => [null, null, null]),
        threshold: null,
        merged: false,
        key: 0,
      };
      const past: Model[] = [],
        future: Model[] = [],
        trace: string[] = [],
        covered = new Set<string>();
      const copy = (value: Model): Model => ({
        rows: value.rows.map((row) => [...row]),
        threshold: value.threshold,
        merged: value.merged,
        key: value.key,
      });
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      for (let step = 0; step < 50; step++) {
        let book = app.getWorkbook(),
          sheet = book.getSheets()[0];
        const action = step < 13 ? [0, 1, 2, 8, 6, 7, 9, 3, 4, 5, 2, 6, 7][step] : random(10);
        const last = model.rows.length - 1;
        const save = () => {
          past.push(copy(model));
          future.length = 0;
        };
        if (action === 0) {
          const row = random(last),
            column = model.key + random(2),
            value = seed + step;
          save();
          await sheet
            .range({ startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 1 })
            .setValues([[value]]);
          model.rows[row][column] = value;
          trace.push('edit');
        } else if (action === 1) {
          save();
          if (model.threshold === null) {
            model.threshold = seed + random(30);
            await sheet.setFilter({
              mode: 'all',
              headerRows: 0,
              criteria: [{ column: model.key, operator: 'gt', value: model.threshold }],
            });
            trace.push('filter');
          } else {
            await sheet.setFilter(null);
            model.threshold = null;
            trace.push('clear filter');
          }
        } else if (action === 2) {
          save();
          const range = sheet.range({
            startRow: last,
            endRow: last + 1,
            startColumn: model.key + 1,
            endColumn: model.key + 3,
          });
          if (model.merged) await range.unmerge();
          else await range.merge();
          model.merged = !model.merged;
          trace.push(model.merged ? 'merge' : 'unmerge');
        } else if (action === 3 && last < 9) {
          if (model.merged && model.threshold !== null) covered.add('insert-filter-merge');
          const row = random(last + 1);
          save();
          await sheet.insertRows(row);
          model.rows.splice(row, 0, Array(model.rows[0].length).fill(null));
          trace.push('insert row');
        } else if (action === 4 && last > 2) {
          if (model.merged && model.threshold !== null) covered.add('delete-filter-merge');
          const row = random(last);
          save();
          await sheet.deleteRows(row);
          model.rows.splice(row, 1);
          trace.push('delete row');
        } else if (action === 5) {
          if (model.merged && model.threshold !== null) covered.add('sort-filter-merge');
          const direction = random(2) ? 'asc' : 'desc';
          save();
          await sheet.sort(
            { startRow: 0, endRow: last, startColumn: 0, endColumn: model.rows[0].length },
            model.key,
            direction,
          );
          const sorted = model.rows
            .slice(0, last)
            .map((row, index) => ({ row, index }))
            .sort((a, b) => {
              const left = a.row[model.key],
                right = b.row[model.key];
              if (left === null && right === null) return a.index - b.index;
              if (left === null) return 1;
              if (right === null) return -1;
              return (direction === 'asc' ? left - right : right - left) || a.index - b.index;
            })
            .map(({ row }) => row);
          model.rows = [...sorted, model.rows[last]];
          trace.push('sort');
        } else if (action === 6 && past.length) {
          const previousKey = model.key;
          future.push(copy(model));
          model = past.pop()!;
          if (!(await book.undo())) return { seed, step, trace, error: 'undo failed' };
          if (previousKey !== model.key && model.merged && model.threshold !== null)
            covered.add('undo-column-filter-merge');
          trace.push('undo');
        } else if (action === 7 && future.length) {
          const previousKey = model.key;
          past.push(copy(model));
          model = future.pop()!;
          if (!(await book.redo())) return { seed, step, trace, error: 'redo failed' };
          if (previousKey !== model.key && model.merged && model.threshold !== null)
            covered.add('redo-column-filter-merge');
          trace.push('redo');
        } else if (action === 8 && model.key < 3) {
          if (model.merged && model.threshold !== null) covered.add('insert-column-filter-merge');
          save();
          await sheet.insertColumns(0);
          for (const row of model.rows) row.unshift(null);
          model.key++;
          trace.push('insert column');
        } else if (action === 9 && model.key > 0) {
          if (model.merged && model.threshold !== null) covered.add('delete-column-filter-merge');
          save();
          await sheet.deleteColumns(0);
          for (const row of model.rows) row.shift();
          model.key--;
          trace.push('delete column');
        } else trace.push('skip');
        covered.add(trace.at(-1)!);
        if (step === 5 || step % 9 === 8) {
          if (model.merged && model.threshold !== null) covered.add('reopen-filter-merge');
          if (model.merged && model.threshold !== null && model.key > 0)
            covered.add('reopen-shifted-filter-merge');
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          trace.push('reopen');
        }
        const actual = await sheet
          .range({
            startRow: 0,
            endRow: model.rows.length,
            startColumn: 0,
            endColumn: model.rows[0].length,
          })
          .getValues();
        const hidden = book.filteredRows(sheet.id);
        const expectedHidden =
          model.threshold === null
            ? []
            : model.rows.flatMap((row, i) =>
                row[model.key] === null || row[model.key]! <= model.threshold! ? [i] : [],
              );
        const metadata = book.sheetData(sheet.id),
          merges = metadata.merges;
        const expectedFilter =
          model.threshold === null
            ? null
            : {
                mode: 'all',
                headerRows: 0,
                criteria: [{ column: model.key, operator: 'gt', value: model.threshold }],
              };
        const expectedMerges = model.merged
          ? [
              {
                startRow: model.rows.length - 1,
                startColumn: model.key + 1,
                endRow: model.rows.length,
                endColumn: model.key + 3,
              },
            ]
          : [];
        if (
          sheet.rowCount !== model.rows.length ||
          sheet.columnCount !== model.rows[0].length ||
          JSON.stringify(actual) !== JSON.stringify(model.rows) ||
          JSON.stringify(hidden) !== JSON.stringify(expectedHidden) ||
          JSON.stringify(metadata.filter) !== JSON.stringify(expectedFilter) ||
          JSON.stringify(merges) !== JSON.stringify(expectedMerges)
        )
          return {
            seed,
            step,
            trace,
            model,
            actual,
            hidden,
            expectedHidden,
            filter: metadata.filter,
            expectedFilter,
            merges,
            expectedMerges,
          };
      }
      for (const action of [
        'edit',
        'filter',
        'clear filter',
        'merge',
        'unmerge',
        'insert row',
        'delete row',
        'sort',
        'undo',
        'redo',
        'insert-filter-merge',
        'delete-filter-merge',
        'sort-filter-merge',
        'reopen-filter-merge',
        'insert column',
        'delete column',
        'insert-column-filter-merge',
        'delete-column-filter-merge',
        'reopen-shifted-filter-merge',
        'undo-column-filter-merge',
        'redo-column-filter-merge',
      ])
        if (!covered.has(action)) return { seed, trace, error: `missing ${action}` };
    }
    return null;
  }, file);
  expect(result).toBeNull();
});
