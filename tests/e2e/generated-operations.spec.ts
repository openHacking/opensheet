import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('seeded mixed edits, axis changes, sorting, history and reopen agree with an independent grid', async ({
  page,
}) => {
  test.setTimeout(120_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const file = JSON.stringify(
    createWorkbookFile({ sheets: [{ name: 'Model', rows: 5, columns: 4 }] }),
  );
  const result = await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    const seeds = [0x1532, 0x7219, 0xbad5];
    for (const seed of seeds) {
      await app.load(JSON.parse(fileJSON));
      let state: (number | null)[][] = Array.from({ length: 5 }, () => Array(4).fill(null));
      let past: (number | null)[][][] = [],
        future: (number | null)[][][] = [];
      let rng = seed;
      const random = (n: number) => {
        rng = (Math.imul(rng, 1664525) + 1013904223) >>> 0;
        return Math.floor((rng / 0x100000000) * n);
      };
      const clone = (grid: (number | null)[][]) => grid.map((row) => row.slice());
      const trace: string[] = [];
      const covered = new Set<string>();
      for (let step = 0; step < 42; step++) {
        let book = app.getWorkbook();
        let sheet = book.getSheets()[0];
        const action = step < 9 ? [0, 3, 5, 7, 8, 9, 7, 8, 9][step] : random(10);
        if (action < 3) {
          const row = random(state.length),
            column = random(state[0].length),
            value = step + seed;
          past.push(clone(state));
          future = [];
          await sheet
            .range({ startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 1 })
            .setValues([[value]]);
          state[row][column] = value;
          trace.push(`set ${row},${column}=${value}`);
        } else if (action === 3 && state.length < 9) {
          const row = random(state.length + 1);
          past.push(clone(state));
          future = [];
          await sheet.insertRows(row);
          state.splice(row, 0, Array(state[0].length).fill(null));
          trace.push(`insert row ${row}`);
        } else if (action === 4 && state.length > 2) {
          const row = random(state.length);
          past.push(clone(state));
          future = [];
          await sheet.deleteRows(row);
          state.splice(row, 1);
          trace.push(`delete row ${row}`);
        } else if (action === 5 && state[0].length < 7) {
          const column = random(state[0].length + 1);
          past.push(clone(state));
          future = [];
          await sheet.insertColumns(column);
          for (const row of state) row.splice(column, 0, null);
          trace.push(`insert column ${column}`);
        } else if (action === 6 && state[0].length > 2) {
          const column = random(state[0].length);
          past.push(clone(state));
          future = [];
          await sheet.deleteColumns(column);
          for (const row of state) row.splice(column, 1);
          trace.push(`delete column ${column}`);
        } else if (action === 7 && past.length) {
          future.push(clone(state));
          state = past.pop()!;
          if (!(await book.undo())) return { seed, step, trace, error: 'undo returned false' };
          trace.push('undo');
        } else if (action === 8 && future.length) {
          past.push(clone(state));
          state = future.pop()!;
          if (!(await book.redo())) return { seed, step, trace, error: 'redo returned false' };
          trace.push('redo');
        } else if (action === 9) {
          const column = random(state[0].length);
          const second = (column + 1 + random(state[0].length - 1)) % state[0].length;
          const keys = [
            { column, direction: random(2) === 0 ? ('asc' as const) : ('desc' as const) },
            { column: second, direction: random(2) === 0 ? ('asc' as const) : ('desc' as const) },
          ];
          const activeKeys =
            step === 5 || (step !== 8 && random(2) === 0) ? keys : keys.slice(0, 1);
          covered.add(activeKeys.length === 2 ? 'multi-key sort' : 'single-key sort');
          past.push(clone(state));
          future = [];
          await sheet.sort(
            { startRow: 0, startColumn: 0, endRow: state.length, endColumn: state[0].length },
            activeKeys,
          );
          state = state
            .map((row, index) => ({ row, index }))
            .sort((a, b) => {
              for (const key of activeKeys) {
                const left = a.row[key.column],
                  right = b.row[key.column];
                if (left === null && right === null) continue;
                if (left === null) return 1;
                if (right === null) return -1;
                const comparison = key.direction === 'asc' ? left - right : right - left;
                if (comparison) return comparison;
              }
              return a.index - b.index;
            })
            .map(({ row }) => row);
          trace.push(`sort ${activeKeys.map((key) => `${key.column}:${key.direction}`).join(',')}`);
        } else trace.push('skip');
        const operation = trace[trace.length - 1].split(' ').slice(0, 2).join(' ');
        covered.add(operation.startsWith('sort ') ? 'sort' : operation);
        if (step === 3 || step === 6 || step % 11 === 10) {
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          trace.push('reopen');
        }
        const actual = await sheet
          .range({
            startRow: 0,
            startColumn: 0,
            endRow: state.length,
            endColumn: state[0].length,
          })
          .getValues();
        if (
          sheet.rowCount !== state.length ||
          sheet.columnCount !== state[0].length ||
          JSON.stringify(actual) !== JSON.stringify(state)
        )
          return {
            seed,
            step,
            trace,
            expected: state,
            actual,
            rows: sheet.rowCount,
            columns: sheet.columnCount,
          };
      }
      for (const operation of [
        'undo',
        'redo',
        'sort',
        'single-key sort',
        'multi-key sort',
        'insert row',
        'delete row',
        'insert column',
        'delete column',
      ]) {
        if (!covered.has(operation)) return { seed, trace, error: `missing ${operation}` };
      }
    }
    return null;
  }, file);
  expect(result).toBeNull();
});
