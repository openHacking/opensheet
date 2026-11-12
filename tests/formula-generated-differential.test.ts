import { expect, it } from 'vitest';
import fc from 'fast-check';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import type { CellInput, CellValue } from '../packages/core/src/types.js';

it('matches pure and asynchronous evaluation over replayable mixed scalar formulas', async () => {
  const value = fc.oneof(
    fc.integer({ min: -100, max: 100 }),
    fc.constantFrom('', '7', '-2.5', 'abc', '😀', '0x10'),
    fc.boolean(),
    fc.constant(null),
    fc.constant({ error: '#DIV/0!' }),
  );
  const base = fc.constantFrom(
    'A1+B1',
    'A1-B1',
    'A1*B1',
    'A1/B1',
    'A1&B1',
    'A1=B1',
    'A1<B1',
    'IF(A1,B1,C1)',
    'IFERROR(A1,B1)',
    'SUM(A1:C1)',
    'AVERAGE(A1:C1)',
    'MIN(A1:C1)',
    'MAX(A1:C1)',
    'COUNT(A1:C1)',
    'COUNTA(A1:C1)',
    'ROUND(A1,B1)',
    'LEFT(A1,B1)',
    'MID(A1,B1,C1)',
    'CONCAT(A1:C1)',
  );
  const expression = fc
    .tuple(base, fc.array(fc.integer({ min: 0, max: 5 }), { maxLength: 4 }))
    .map(([initial, wrappers]) =>
      wrappers.reduce<string>((current, wrapper) => {
        switch (wrapper) {
          case 0:
            return `IF(TRUE,${current},1/0)`;
          case 1:
            return `IFERROR(${current},B1)`;
          case 2:
            return `SUM(${current},C1)`;
          case 3:
            return `-(${current})`;
          case 4:
            return `(${current})+A1`;
          default:
            return `IF(A1,${current},B1)`;
        }
      }, initial),
    );
  const input = (value: CellValue): CellInput => {
    if (value === null) return { type: 'blank' };
    if (typeof value === 'number') return { type: 'number', value };
    if (typeof value === 'boolean') return { type: 'boolean', value };
    if (typeof value === 'string') return { type: 'string', value };
    return { type: 'error', code: value.error };
  };
  await fc.assert(
    fc.asyncProperty(
      expression,
      fc.tuple(value, value, value),
      fc.boolean(),
      async (formula, values, cached) => {
        const expected = evaluateFormula(
          parseFormula(formula),
          (ref) => values[ref.column] ?? null,
        );
        const calculation = new AsyncCalculation(
          async (_sheet, row, column) => ({
            rowId: row,
            columnId: column,
            input:
              column === 3
                ? { type: 'formula', expression: formula }
                : input(values[column] ?? null),
          }),
          () => undefined,
          cached ? 1024 * 1024 : 0,
        );
        expect(await calculation.value('sheet', 0, 3)).toEqual(expected);
      },
    ),
    { seed: 20261112, numRuns: 5000 },
  );
});
