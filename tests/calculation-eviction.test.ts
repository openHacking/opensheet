import { expect, it, vi } from 'vitest';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';

it('reuses overlapping scalar reads only while a read session remains current', async () => {
  let sourceReads = 0;
  let source = 1;
  const calculation = new AsyncCalculation(
    async (sheet, row, column) => {
      if (sheet === 'source') {
        sourceReads++;
        return { rowId: row, columnId: column, input: { type: 'number', value: source } };
      }
      return {
        rowId: row,
        columnId: column,
        input: { type: 'formula', expression: `SUM(Source!$A$1:$A$32)+${row}` },
      };
    },
    (name) => (name === 'Source' ? 'source' : undefined),
  );
  const read = calculation.readSession();
  for (let row = 0; row < 128; row++) expect(await read('formula', row, 0)).toBe(32 + row);
  expect(sourceReads).toBe(32);
  source = 2;
  calculation.invalidate([
    { sheetId: 'source', startRow: 0, endRow: 32, startColumn: 0, endColumn: 1 },
  ]);
  for (let row = 0; row < 128; row++) expect(await read('formula', row, 0)).toBe(64 + row);
  expect(sourceReads).toBe(64);
});

it.each([0, 4 * 1024 * 1024])(
  'does not retain stale scalar results when invalidation races an in-flight read with budget %i',
  async (budget) => {
    let release!: () => void;
    let started!: () => void;
    const paused = new Promise<void>((resolve) => {
      started = resolve;
    });
    const resume = new Promise<void>((resolve) => {
      release = resolve;
    });
    let reads = 0;
    const calculation = new AsyncCalculation(
      async () => {
        const value = ++reads;
        if (value === 1) {
          started();
          await resume;
        }
        return { rowId: 0, columnId: 0, input: { type: 'number', value } };
      },
      () => undefined,
      budget,
    );
    const read = calculation.readSession(16 * 1024);
    const first = read('s', 0, 0);
    await paused;
    calculation.invalidate([
      { sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    ]);
    release();
    expect(await first).toBe(1);
    expect(await read('s', 0, 0)).toBe(2);
    expect(await read('s', 0, 0)).toBe(2);
    expect(reads).toBe(2);
  },
);

it('does not memoize scalar references when the calculation cache budget is zero', async () => {
  let reads = 0;
  const calculation = new AsyncCalculation(
    async () => ({ rowId: 0, columnId: 0, input: { type: 'number', value: ++reads } }),
    () => undefined,
    0,
  );
  const read = calculation.readSession();
  expect(await read('s', 0, 0)).toBe(1);
  expect(await read('s', 0, 0)).toBe(2);
});

it('reuses bounded read scratch across uncached formulas and invalidates on edits', async () => {
  let reads = 0;
  let source = 1;
  const calculation = new AsyncCalculation(
    async (sheet, row, column) => {
      if (sheet === 'source') {
        reads++;
        return { rowId: row, columnId: column, input: { type: 'number', value: source } };
      }
      return {
        rowId: row,
        columnId: column,
        input: { type: 'formula', expression: `SUM(Source!$A$1:$A$32)+${row}` },
      };
    },
    (name) => (name === 'Source' ? 'source' : undefined),
    0,
  );
  const read = calculation.readSession(16 * 1024);
  for (let row = 0; row < 128; row++) expect(await read('formula', row, 0)).toBe(32 + row);
  expect(reads).toBe(32);
  expect(calculation.cacheBytes).toBe(0);
  source = 2;
  calculation.invalidate([
    { sheetId: 'source', startRow: 0, endRow: 32, startColumn: 0, endColumn: 1 },
  ]);
  for (let row = 0; row < 128; row++) expect(await read('formula', row, 0)).toBe(64 + row);
  expect(reads).toBe(64);
  expect(calculation.cacheBytes).toBe(0);
  expect(() => calculation.readSession(65537)).toThrow();
});

it('bounds the read-session scalar memo and skips values larger than its budget', async () => {
  let reads = 0;
  const calculation = new AsyncCalculation(
    async (_sheet, row, column) => {
      reads++;
      return {
        rowId: row,
        columnId: column,
        input:
          row === 999
            ? { type: 'string', value: 'x'.repeat(32768) }
            : { type: 'number', value: row },
      };
    },
    () => undefined,
  );
  const read = calculation.readSession();
  for (let row = 0; row < 300; row++) expect(await read('s', row, 0)).toBe(row);
  expect(reads).toBe(300);
  expect(await read('s', 0, 0)).toBe(0);
  expect(reads).toBe(301);
  expect(await read('s', 999, 0)).toHaveLength(32768);
  expect(await read('s', 999, 0)).toHaveLength(32768);
  expect(reads).toBe(303);
});

it('invalidates a cached dependent when its intermediate formula was evicted', async () => {
  let source = 1;
  const read = vi.fn(async (_sheet: string, _row: number, column: number) => ({
    rowId: 0,
    columnId: column,
    input:
      column === 0
        ? { type: 'number' as const, value: source }
        : { type: 'formula' as const, expression: column === 1 ? 'A1+1' : 'B1+1' },
  }));
  const calculation = new AsyncCalculation(read, () => undefined, 400);
  expect(await calculation.value('s', 0, 2)).toBe(3);
  read.mockClear();
  expect(await calculation.value('s', 0, 2)).toBe(3);
  expect(read).toHaveBeenCalled();
  source = 7;
  calculation.invalidate([{ sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]);
  expect(await calculation.value('s', 0, 2)).toBe(9);
});

it('evicts downstream cached formulas when an older intermediate is displaced', async () => {
  let source = 1;
  const read = vi.fn(async (_sheet: string, _row: number, column: number) => ({
    rowId: 0,
    columnId: column,
    input:
      column === 0
        ? { type: 'number' as const, value: source }
        : {
            type: 'formula' as const,
            expression: column === 1 ? 'A1+1' : column === 2 ? 'B1+1' : '42',
          },
  }));
  const calculation = new AsyncCalculation(read, () => undefined, 950);
  expect(await calculation.value('s', 0, 2)).toBe(3);
  read.mockClear();
  expect(await calculation.value('s', 0, 2)).toBe(3);
  expect(read).not.toHaveBeenCalled();
  await calculation.value('s', 0, 3);
  read.mockClear();
  expect(await calculation.value('s', 0, 2)).toBe(3);
  expect(read).toHaveBeenCalled();
  source = 7;
  calculation.invalidate([{ sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]);
  expect(await calculation.value('s', 0, 2)).toBe(9);
});

it.each(['invalidate', 'clear'] as const)(
  'does not cache an in-flight result after %s even when the cache is empty',
  async (operation) => {
    let source = 1;
    let release!: () => void;
    let started!: () => void;
    const paused = new Promise<void>((resolve) => {
      started = resolve;
    });
    const resume = new Promise<void>((resolve) => {
      release = resolve;
    });
    let first = true;
    const calculation = new AsyncCalculation(
      async (_sheet, _row, column) => {
        if (column === 1)
          return { rowId: 0, columnId: column, input: { type: 'formula', expression: 'A1+1' } };
        const value = source;
        if (first) {
          first = false;
          started();
          await resume;
        }
        return { rowId: 0, columnId: column, input: { type: 'number', value } };
      },
      () => undefined,
    );
    const pending = calculation.value('s', 0, 1);
    await paused;
    source = 7;
    if (operation === 'clear') calculation.clear();
    else
      calculation.invalidate([
        { sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
      ]);
    release();
    expect(await pending).toBe(2);
    expect(await calculation.value('s', 0, 1)).toBe(8);
  },
);

it('invalidates a long cached chain without recursive eviction or stale outputs', async () => {
  let source = 1;
  const calculation = new AsyncCalculation(
    async (_sheet, row) => ({
      rowId: row,
      columnId: 0,
      input:
        row === 0
          ? { type: 'number', value: source }
          : { type: 'formula', expression: `A${row}+1` },
    }),
    () => undefined,
    1024 * 1024,
  );
  for (let row = 1; row <= 500; row++)
    expect(await calculation.value('s', row, 0)).toEqual(
      row <= 128 ? row + 1 : { error: '#LIMIT!' },
    );
  expect(calculation.cacheBytes).toBeLessThanOrEqual(1024 * 1024);
  source = 5;
  calculation.invalidate([{ sheetId: 's', startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }]);
  for (let row = 1; row <= 500; row++)
    expect(await calculation.value('s', row, 0)).toEqual(
      row <= 128 ? row + 5 : { error: '#LIMIT!' },
    );
  expect(calculation.cacheBytes).toBeLessThanOrEqual(1024 * 1024);
});

it('keeps the depth limit stable after formula cache eviction', async () => {
  const budget = 128 * 1024;
  const calculation = new AsyncCalculation(
    async (sheet, row, column) => ({
      rowId: row,
      columnId: column,
      input:
        sheet === 'filler'
          ? { type: 'formula' as const, expression: `${row}+1` }
          : row === 0
            ? { type: 'number' as const, value: 1 }
            : { type: 'formula' as const, expression: `A${row}+1` },
    }),
    () => undefined,
    budget,
  );
  expect(await calculation.value('chain', 129, 0)).toEqual({ error: '#LIMIT!' });
  expect(await calculation.value('chain', 128, 0)).toBe(129);
  expect(await calculation.value('chain', 129, 0)).toEqual({ error: '#LIMIT!' });
  for (let row = 0; row < 2000; row++)
    expect(await calculation.value('filler', row, 0)).toBe(row + 1);
  expect(calculation.cacheBytes).toBeLessThanOrEqual(budget);
  expect(await calculation.value('chain', 129, 0)).toEqual({ error: '#LIMIT!' });
});
