import { expect, it } from 'vitest';
import { createWorkbook } from '@opensheetjs/core';
import { offsetFormula } from '@opensheetjs/formula';
import {
  createTileOffsetter,
  tileCell,
  validatePasteTile,
  type PasteTile,
} from '../packages/core/src/paste-tile.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it('rejects cumulative relocation growth beyond the clipboard budget atomically', () => {
  const count = 1100;
  const budget = 32 * 1024 * 1024;
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: count, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  const expression = 'A1+A1';
  const item = {
    input: { type: 'formula' as const, expression },
    source: { row: 0, column: 0 },
    note: '',
  };
  const overhead = JSON.stringify(item).length + 32;
  item.note = 'x'.repeat(Math.floor((budget - 1) / count) - overhead);
  const tile: PasteTile = {
    sheetId: sheet.id,
    range: { startRow: 0, endRow: count, startColumn: 0, endColumn: 1 },
    rows: 1,
    columns: 1,
    cells: [item],
  };
  const base = validatePasteTile(sheet, tile);
  const growth = Array.from({ length: count }, (_, row) =>
    Math.max(0, offsetFormula(expression, row, 0).length - expression.length),
  );
  expect(base).toBeLessThanOrEqual(budget);
  expect(base + growth.reduce((a, b) => a + b, 0)).toBeGreaterThan(budget);
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, { type: 'core.cells.pasteTile', payload: tile }),
  ).toThrowError(
    expect.objectContaining({
      code: 'LIMIT_EXCEEDED',
      message: 'Expanded clipboard content is too large',
    }),
  );
  expect(snapshot).toEqual(before);
});

it('does not mutate the snapshot when a later relocated formula exceeds the text limit', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 1, columns: 2 }] });
  const sheet = snapshot.sheets[0];
  const expression = `A1&"${'x'.repeat(32763)}"`;
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.cells.pasteTile',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
        rows: 1,
        columns: 2,
        cells: [
          { input: { type: 'number', value: 42 } },
          { input: { type: 'formula', expression }, source: { row: 0, column: 0 } },
        ],
      },
    }),
  ).toThrowError(expect.objectContaining({ code: 'LIMIT_EXCEEDED' }));
  expect(snapshot).toEqual(before);
});

it('rejects formula relocation that would exceed the cell text limit', () => {
  const expression = `A1&"${'x'.repeat(32763)}"`;
  expect(expression).toHaveLength(32768);
  const tile: PasteTile = {
    sheetId: 'sheet',
    range: { startRow: 1, endRow: 3, startColumn: 0, endColumn: 1 },
    rows: 1,
    columns: 1,
    cells: [{ input: { type: 'formula', expression }, source: { row: 0, column: 0 } }],
  };
  expect(() => tileCell(tile, 1, 0, () => {})).toThrowError(
    expect.objectContaining({ code: 'LIMIT_EXCEEDED' }),
  );
});

it('rejects malformed tiles atomically and preserves unrelated data', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 100, columns: 26 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[7]]);
    const revision = book.revision;
    const command = {
      type: 'core.cells.pasteTile',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
        rows: 1,
        columns: 2,
        cells: [{ input: { type: 'number', value: 5 } }],
      },
    };
    await expect(book.execute(command)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1:B2').getValues()).toEqual([
      [7, null],
      [null, null],
    ]);
  } finally {
    await book.deleteStorage();
  }
});

it('preserves relocation results after the per-command formula cache is full', () => {
  const offset = createTileOffsetter();
  for (let i = 1; i <= 200; i++) {
    const formula = `A${i}+$B$1`;
    expect(offset(formula, 2, -1)).toBe(offsetFormula(formula, 2, -1));
  }
});
