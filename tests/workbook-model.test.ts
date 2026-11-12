import { appendFileSync, writeFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import fc from 'fast-check';
import { createWorkbook, openWorkbook, type Rect } from '@opensheetjs/core';

// Deliberately independent of reducers, formula parsers, comparators and filter helpers.
type Formula = {
  row: number;
  column: number;
  absRow: boolean;
  absColumn: boolean;
  broken?: boolean;
};
type Cell = number | null | Formula;
type Value = number | null | { error: string };
type State = {
  cells: Cell[][];
  rows: number[];
  columns: number[];
  nextRow: number;
  nextColumn: number;
  merges: Rect[];
  filter: null | {
    mode: 'all' | 'any';
    headerRows: number;
    criteria: { column: number; operator: 'gte' | 'lte'; value: number }[];
  };
};
const clone = <T>(value: T): T => structuredClone(value);
function valueAt(state: State, row: number, column: number, visited = new Set<string>()): Value {
  if (row < 0 || column < 0 || row >= state.rows.length || column >= state.columns.length)
    return { error: '#REF!' };
  const input = state.cells[row][column];
  if (input === null || typeof input === 'number') return input;
  if (input.broken) return { error: '#REF!' };
  const key = `${row}:${column}`;
  if (visited.has(key)) return { error: '#CYCLE!' };
  visited.add(key);
  const result = valueAt(state, input.row, input.column, visited);
  visited.delete(key);
  return result && typeof result === 'object' ? result : (result ?? 0) * 2;
}
function values(state: State) {
  return state.cells.map((row, r) => row.map((_, c) => valueAt(state, r, c)));
}
function hidden(state: State): number[] {
  if (!state.filter) return [];
  const filter = state.filter;
  return state.rows.flatMap((_, row) => {
    if (row < filter.headerRows) return [];
    const matches = filter.criteria.map((c) => {
      const v = valueAt(state, row, c.column);
      return typeof v === 'number' && (c.operator === 'gte' ? v >= c.value : v <= c.value);
    });
    return (filter.mode === 'all' ? matches.every(Boolean) : matches.some(Boolean)) ? [] : [row];
  });
}
function compare(a: Value, b: Value, descending: boolean) {
  if (a === b) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  const rank =
    typeof a === 'number' ? (typeof b === 'number' ? 0 : -1) : typeof b === 'number' ? 1 : 0;
  const x = typeof a === 'number' ? a : a.error,
    y = typeof b === 'number' ? b : b.error;
  return (rank || (x === y ? 0 : x < y ? -1 : 1)) * (descending ? -1 : 1);
}
const operations = [
  'edit',
  'formula',
  'insertRow',
  'deleteRow',
  'insertColumn',
  'deleteColumn',
  'sort',
  'filter',
  'clearFilter',
  'merge',
  'unmerge',
  'undo',
  'redo',
  'reopen',
  'rollback',
  'stale',
] as const;
const action = fc.record({
  op: fc.constantFrom(...operations),
  r: fc.nat(15),
  c: fc.nat(15),
  rr: fc.nat(15),
  cc: fc.nat(15),
  n: fc.integer({ min: -10, max: 10 }),
  flag: fc.boolean(),
  other: fc.boolean(),
  blank: fc.boolean(),
});
const reportPath = process.env.OPENSHEET_MODEL_REPORT;
if (reportPath) writeFileSync(reportPath, '', { flag: 'wx' });
const seeds = process.env.OPENSHEET_MODEL_SEED
  ? [Number(process.env.OPENSHEET_MODEL_SEED)]
  : [20261023, 73013, 991];
it.each(seeds)(
  'matches an independent workbook model through generated commands (seed %s)',
  async (seed) => {
    const coverage = new Map<string, number>();
    await fc.assert(
      fc.asyncProperty(fc.array(action, { minLength: 40, maxLength: 70 }), async (actions) => {
        let book = await createWorkbook({ sheets: [{ name: 'Model', rows: 4, columns: 3 }] });
        let state: State = {
          cells: Array.from({ length: 4 }, () => [null, null, null]),
          rows: [0, 1, 2, 3],
          columns: [0, 1, 2],
          nextRow: 4,
          nextColumn: 3,
          merges: [],
          filter: null,
        };
        const past: State[] = [],
          future: State[] = [];
        let revision = 0;
        try {
          for (const a of actions) {
            const sheet = book.getSheets()[0],
              row = a.r % state.rows.length,
              column = a.c % state.columns.length;
            const range = {
              startRow: 0,
              endRow: state.rows.length,
              startColumn: 0,
              endColumn: state.columns.length,
            };
            const next = clone(state);
            let changes = false;
            const covered = state.merges.some(
              (m) =>
                row >= m.startRow &&
                row < m.endRow &&
                column >= m.startColumn &&
                column < m.endColumn &&
                (row !== m.startRow || column !== m.startColumn),
            );
            if (a.op === 'edit' || a.op === 'formula') {
              const ref = {
                row: a.rr % state.rows.length,
                column: a.cc % state.columns.length,
                absRow: a.flag,
                absColumn: a.other,
              };
              const operation = () =>
                a.op === 'edit'
                  ? sheet
                      .range({
                        startRow: row,
                        endRow: row + 1,
                        startColumn: column,
                        endColumn: column + 1,
                      })
                      .setValues([[a.blank ? null : a.n]])
                  : sheet
                      .range({
                        startRow: row,
                        endRow: row + 1,
                        startColumn: column,
                        endColumn: column + 1,
                      })
                      .setFormulas([
                        [
                          `${ref.absColumn ? '$' : ''}${String.fromCharCode(65 + ref.column)}${ref.absRow ? '$' : ''}${ref.row + 1}*2`,
                        ],
                      ]);
              if (covered)
                await expect(operation()).rejects.toMatchObject({ code: 'INVALID_RANGE' });
              else {
                await operation();
                next.cells[row][column] = a.op === 'edit' ? (a.blank ? null : a.n) : ref;
                changes = true;
              }
            } else if (['insertRow', 'deleteRow', 'insertColumn', 'deleteColumn'].includes(a.op)) {
              const isRow = a.op.endsWith('Row'),
                insert = a.op.startsWith('insert');
              const count = isRow ? state.rows.length : state.columns.length;
              const at = (isRow ? a.r : a.c) % (count + Number(insert));
              const start = isRow ? 'startRow' : 'startColumn',
                end = isRow ? 'endRow' : 'endColumn';
              const intersects = state.merges.some(
                (m) => at < m[end] && at + Number(!insert) > m[start],
              );
              if ((insert && count === 8) || (!insert && count === 1)) continue;
              const operation = () =>
                isRow
                  ? insert
                    ? sheet.insertRows(at)
                    : sheet.deleteRows(at)
                  : insert
                    ? sheet.insertColumns(at)
                    : sheet.deleteColumns(at);
              if (intersects)
                await expect(operation()).rejects.toMatchObject({ code: 'INVALID_RANGE' });
              else {
                await operation();
                changes = true;
                for (const cells of next.cells)
                  for (const cell of cells)
                    if (cell && typeof cell === 'object' && !cell.broken) {
                      const field = isRow ? 'row' : 'column';
                      if (!insert && cell[field] === at) cell.broken = true;
                      else if (cell[field] >= at) cell[field] += insert ? 1 : -1;
                    }
                if (isRow) {
                  next.cells.splice(
                    at,
                    Number(!insert),
                    ...(insert ? [Array<Cell>(state.columns.length).fill(null)] : []),
                  );
                  next.rows.splice(at, Number(!insert), ...(insert ? [next.nextRow++] : []));
                  if (next.filter && at < next.filter.headerRows)
                    next.filter.headerRows += insert ? 1 : -1;
                } else {
                  for (const cells of next.cells)
                    cells.splice(at, Number(!insert), ...(insert ? [null] : []));
                  next.columns.splice(at, Number(!insert), ...(insert ? [next.nextColumn++] : []));
                  if (next.filter) {
                    next.filter.criteria = next.filter.criteria.filter(
                      (c) => insert || c.column !== at,
                    );
                    for (const c of next.filter.criteria)
                      if (c.column >= at) c.column += insert ? 1 : -1;
                    if (!next.filter.criteria.length) next.filter = null;
                  }
                }
                for (const m of next.merges)
                  if (m[start] >= at) {
                    m[start] += insert ? 1 : -1;
                    m[end] += insert ? 1 : -1;
                  }
              }
            } else if (a.op === 'sort') {
              const secondary = (column + 1) % state.columns.length;
              const keys = [
                { column, direction: a.flag ? ('desc' as const) : ('asc' as const) },
                ...(secondary === column
                  ? []
                  : [
                      {
                        column: secondary,
                        direction: a.other ? ('desc' as const) : ('asc' as const),
                      },
                    ]),
              ];
              if (state.merges.length)
                await expect(sheet.sort(range, keys)).rejects.toMatchObject({
                  code: 'INVALID_RANGE',
                });
              else {
                const before = values(state);
                const order = state.rows
                  .map((_, i) => i)
                  .sort(
                    (x, y) =>
                      compare(before[x][column], before[y][column], a.flag) ||
                      (secondary === column
                        ? 0
                        : compare(before[x][secondary], before[y][secondary], a.other)) ||
                      x - y,
                  );
                next.cells = order.map((source, destination) =>
                  clone(state.cells[source]).map((cell) => {
                    if (cell && typeof cell === 'object' && !cell.broken && !cell.absRow) {
                      cell.row += destination - source;
                      if (cell.row < 0) cell.broken = true;
                    }
                    return cell;
                  }),
                );
                await sheet.sort(range, keys);
                changes = true;
              }
            } else if (a.op === 'filter' || a.op === 'clearFilter') {
              next.filter =
                a.op === 'clearFilter'
                  ? null
                  : {
                      mode: a.flag ? 'all' : 'any',
                      headerRows: Number(a.other),
                      criteria: [
                        { column, operator: 'gte', value: a.n },
                        {
                          column: (column + 1) % state.columns.length,
                          operator: 'lte',
                          value: a.n + 5,
                        },
                      ],
                    };
              await sheet.setFilter(next.filter);
              changes = true;
            } else if (a.op === 'merge') {
              if (state.columns.length < 2 || state.merges.length) continue;
              const merge = { startRow: row, endRow: row + 1, startColumn: 0, endColumn: 2 };
              await book.execute({
                type: 'core.cells.merge',
                payload: { sheetId: sheet.id, range: merge, discardCoveredValues: true },
              });
              next.merges.push(merge);
              next.cells[row][1] = null;
              changes = true;
            } else if (a.op === 'unmerge') {
              await sheet.range(range).unmerge();
              next.merges = [];
              changes = true;
            } else if (a.op === 'undo' || a.op === 'redo') {
              const from = a.op === 'undo' ? past : future,
                to = a.op === 'undo' ? future : past;
              await (a.op === 'undo' ? book.undo() : book.redo());
              if (from.length) {
                to.push(clone(state));
                state = from.pop()!;
                revision++;
              }
            } else if (a.op === 'reopen') {
              const id = book.id;
              await book.close();
              book = await openWorkbook(id, { cacheBytes: a.flag ? 0 : 65536 });
            } else if (a.op === 'rollback') {
              await expect(
                book.transaction({}, async (tx) => {
                  await tx
                    .getSheets()[0]
                    .range('A1')
                    .setValues([[999]]);
                  await tx.execute({
                    type: 'core.sheet.filter',
                    payload: { sheetId: sheet.id, filter: { column: 999, query: 'bad' } },
                  });
                }),
              ).rejects.toThrow();
            } else {
              await expect(
                book.execute({
                  protocolVersion: 1,
                  commandId: crypto.randomUUID(),
                  workbookId: book.id,
                  baseRevision: revision + 1,
                  type: 'core.sheet.filter',
                  payload: { sheetId: sheet.id, filter: null },
                }),
              ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
            }
            coverage.set(a.op, (coverage.get(a.op) ?? 0) + 1);
            if (changes) {
              past.push(clone(state));
              future.length = 0;
              state = next;
              revision++;
            }
            const current = book.getSheets()[0],
              metadata = book.sheetData(current.id);
            expect(
              await current
                .range({
                  startRow: 0,
                  endRow: state.rows.length,
                  startColumn: 0,
                  endColumn: state.columns.length,
                })
                .getValues(),
            ).toEqual(values(state));
            expect(metadata.rowOrder).toEqual(state.rows);
            expect(metadata.columnOrder).toEqual(state.columns);
            expect(metadata.merges).toEqual(state.merges);
            expect(metadata.filter).toEqual(state.filter);
            expect(book.filteredRows(current.id)).toEqual(hidden(state));
            expect(book.revision).toBe(revision);
            expect(book.canUndo).toBe(past.length > 0);
            expect(book.canRedo).toBe(future.length > 0);
          }
        } finally {
          try {
            await book.deleteStorage();
          } finally {
            await book.close();
          }
        }
      }),
      {
        seed,
        numRuns: Number(process.env.OPENSHEET_MODEL_RUNS ?? 12),
        ...(process.env.OPENSHEET_MODEL_PATH ? { path: process.env.OPENSHEET_MODEL_PATH } : {}),
      },
    );
    if (!process.env.OPENSHEET_MODEL_PATH)
      expect([...coverage.keys()].sort()).toEqual([...operations].sort());
    if (reportPath)
      appendFileSync(
        reportPath,
        JSON.stringify({
          seed,
          runs: Number(process.env.OPENSHEET_MODEL_RUNS ?? 12),
          checkedTransitions: [...coverage.values()].reduce((sum, n) => sum + n, 0),
          operations: Object.fromEntries(coverage),
        }) + '\n',
      );
  },
  Number(process.env.OPENSHEET_MODEL_TIMEOUT_MS ?? 60000),
);
