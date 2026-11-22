import { expect, it } from 'vitest';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';

it('yields for repeated references even when budget skips multiples of 1024', async () => {
  const group = `SUM(${Array.from({ length: 254 }, () => 'A1').join(',')},0)`;
  const expression = `SUM(${Array.from({ length: 5 }, () => group).join(',')})`;
  const calculation = new AsyncCalculation(
    async (_sheet, _row, column) => ({
      rowId: 0,
      columnId: column,
      input: column === 0 ? { type: 'number', value: 1 } : { type: 'formula', expression },
    }),
    () => undefined,
    0,
  );
  let checkpoints = 0;
  const read = calculation.readSession(16 * 1024, () => checkpoints++);
  expect(await read('sheet', 0, 1)).toBe(1270);
  expect(checkpoints).toBeGreaterThan(0);
});

it('shares the reference counter across uncached referenced formulas', async () => {
  const child = `SUM(${Array.from({ length: 254 }, () => 'A1').join(',')},0)`;
  const expression = 'SUM(C1,C1,C1,C1,C1)';
  const calculation = new AsyncCalculation(
    async (_sheet, _row, column) => ({
      rowId: 0,
      columnId: column,
      input:
        column === 0
          ? { type: 'number', value: 1 }
          : { type: 'formula', expression: column === 1 ? expression : child },
    }),
    () => undefined,
    0,
  );
  let checkpoints = 0;
  const read = calculation.readSession(16 * 1024, () => checkpoints++);
  expect(await read('sheet', 0, 1)).toBe(1270);
  expect(checkpoints).toBe(1);
});

it.each([1023, 1024, 2048])('yields at exact reference-visit boundaries (%i)', async (count) => {
  const group = `SUM(${Array.from({ length: 254 }, () => 'A1').join(',')},0)`;
  const groups = Math.floor(count / 254);
  const tail = count % 254;
  const expression =
    count === 2048
      ? 'SUM(A1:A1024,A1:A1024)'
      : `SUM(${[...Array.from({ length: groups }, () => group), ...Array.from({ length: tail }, () => 'A1')].join(',')})`;
  const calculation = new AsyncCalculation(
    async (_sheet, _row, column) => ({
      rowId: 0,
      columnId: column,
      input: column === 0 ? { type: 'number', value: 1 } : { type: 'formula', expression },
    }),
    () => undefined,
    0,
  );
  let checkpoints = 0;
  expect(await calculation.readSession(16 * 1024, () => checkpoints++)('sheet', 0, 1)).toBe(count);
  expect(checkpoints).toBe(Math.floor(count / 1024));
});
