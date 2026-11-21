import { expect, it } from 'vitest';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import { FormulaCalculation } from '../packages/core/src/calculation.js';
import { createSheet } from '../packages/core/src/model.js';

it.each([0, 1024, 4 * 1024 * 1024])(
  'keeps lazy, cross-sheet and boundary results independent of cache with budget %i',
  async (budget) => {
    const chain = createSheet('Chain', 131, 1);
    const branch = createSheet('Branch', 3, 1);
    chain.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } };
    for (let row = 1; row < 131; row++)
      chain.cells[`${row}:0`] = {
        rowId: row,
        columnId: 0,
        input: { type: 'formula', expression: `A${row}+1` },
      };
    for (const [row, expression] of [
      [0, 'IF(FALSE,Chain!A130,42)'],
      [1, 'IF(TRUE,Chain!A130,42)'],
      [2, 'Chain!A128'],
    ] as const)
      branch.cells[`${row}:0`] = {
        rowId: row,
        columnId: 0,
        input: { type: 'formula', expression },
      };
    const sheets = new Map([
      [chain.id, chain],
      [branch.id, branch],
    ]);
    const find = (name: string) => (name === 'Chain' ? chain.id : undefined);
    const sync = new FormulaCalculation((id) => sheets.get(id)!, find);
    const asyncCalc = new AsyncCalculation(
      async (id, row, column) => sheets.get(id)?.cells[`${row}:${column}`],
      find,
      budget,
    );
    for (let round = 0; round < 2; round++) {
      for (const row of [1, 64, 127, 128]) {
        expect(sync.value(chain.id, row, 0)).toBe(row + 1);
        expect(await asyncCalc.value(chain.id, row, 0)).toBe(row + 1);
      }
      for (const [id, row, expected] of [
        [chain.id, 129, { error: '#LIMIT!' }],
        [branch.id, 0, 42],
        [branch.id, 1, { error: '#LIMIT!' }],
        [branch.id, 2, 128],
      ] as const) {
        expect(sync.value(id, row, 0)).toEqual(expected);
        expect(await asyncCalc.value(id, row, 0)).toEqual(expected);
      }
    }
  },
);

it('bounds snapshot formula cache and recomputes evicted results', () => {
  const sheet = createSheet('Cache', 35_000, 1);
  for (let row = 0; row < sheet.rowOrder.length; row++)
    sheet.cells[`${row}:0`] = {
      rowId: row,
      columnId: 0,
      input: { type: 'formula', expression: `${row}+1` },
    };
  let reads = 0;
  sheet.cells = new Proxy(sheet.cells, {
    get(target, property, receiver) {
      if (property === '0:0') reads++;
      return Reflect.get(target, property, receiver);
    },
  });
  const calc = new FormulaCalculation(
    () => sheet,
    () => undefined,
  );
  for (let row = 0; row < sheet.rowOrder.length; row++) {
    expect(calc.value(sheet.id, row, 0)).toBe(row + 1);
    expect(calc.cacheBytes).toBeLessThanOrEqual(4 * 1024 * 1024);
  }
  expect(reads).toBe(1);
  expect(calc.value(sheet.id, 0, 0)).toBe(1);
  expect(reads).toBe(2);
  expect(calc.value(sheet.id, 0, 0)).toBe(1);
  expect(reads).toBe(2);
  calc.clear();
  expect(calc.cacheBytes).toBe(0);
});

it('agrees across cache budgets and read orders on mixed-depth dependency graphs', async () => {
  const sheet = createSheet('Graph', 190, 2);
  for (let row = 0; row < 190; row++) {
    sheet.cells[`${row}:0`] = {
      rowId: row,
      columnId: 0,
      input:
        row === 0 ? { type: 'number', value: 2 } : { type: 'formula', expression: `A${row}+1` },
    };
    sheet.cells[`${row}:1`] = {
      rowId: row,
      columnId: 1,
      input: {
        type: 'formula',
        expression: `IF(${row % 3 === 0 ? 'FALSE' : 'TRUE'},A${row + 1},7)`,
      },
    };
  }
  const read = async (_id: string, row: number, column: number) => sheet.cells[`${row}:${column}`];
  const cold = new AsyncCalculation(read, () => undefined, 0);
  const warm = new AsyncCalculation(read, () => undefined, 4 * 1024 * 1024);
  const sync = new FormulaCalculation(
    () => sheet,
    () => undefined,
  );
  for (let row = 0; row < 190; row++) await warm.value(sheet.id, row, 0);
  for (let row = 189; row >= 0; row--) {
    for (const column of [0, 1]) {
      const expected = await cold.value(sheet.id, row, column);
      expect(await warm.value(sheet.id, row, column)).toEqual(expected);
      expect(sync.value(sheet.id, row, column)).toEqual(expected);
    }
  }
});
