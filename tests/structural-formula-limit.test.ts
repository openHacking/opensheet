import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it('skips formulas that are removed but still rejects surviving unparseable formulas atomically', () => {
  for (const axis of ['row', 'column'] as const) {
    const snapshot = createSnapshot({ sheets: [{ name: 'S', rows: 3, columns: 3 }] });
    const sheet = snapshot.sheets[0];
    sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'formula', expression: 'A1+' } };
    const survivor = { rowId: 1, columnId: 1 };
    sheet.cells['1:1'] = { ...survivor, input: { type: 'formula', expression: 'A2*2' } };
    const command = {
      type: 'core.axis.delete' as const,
      payload: { sheetId: sheet.id, axis, index: 0, count: 1 },
    };
    reduceCommand(snapshot, command);
    expect(sheet.cells['0:0']).toBeUndefined();
    expect(sheet.cells['1:1'].input.type).toBe('formula');

    sheet.cells['1:1'].input = { type: 'formula', expression: 'A1+' };
    const before = structuredClone(snapshot);
    expect(() =>
      reduceCommand(snapshot, { ...command, payload: { ...command.payload, index: 1 } }),
    ).toThrowError(expect.objectContaining({ code: 'UNSUPPORTED_FEATURE' }));
    expect(snapshot).toEqual(before);
  }
});

it('can remove a sheet containing an unparseable formula', () => {
  const snapshot = createSnapshot({
    sheets: [
      { name: 'Gone', rows: 2, columns: 2 },
      { name: 'Keep', rows: 2, columns: 2 },
    ],
  });
  const [gone, keep] = snapshot.sheets;
  gone.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'formula', expression: 'A1+' } };
  keep.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'formula', expression: 'Gone!A1' } };
  reduceCommand(snapshot, { type: 'core.sheet.remove', payload: { sheetId: gone.id } });
  expect(snapshot.sheets.map((sheet) => sheet.id)).toEqual([keep.id]);
  expect(keep.cells['0:0'].input).toEqual({ type: 'formula', expression: '#REF!' });
});

it('rejects sheet rename when a later rewritten formula exceeds the limit, without mutation', () => {
  const snapshot = createSnapshot({
    sheets: [
      { name: 'S', rows: 2, columns: 2 },
      { name: 'Other', rows: 2, columns: 2 },
    ],
  });
  const [sheet, other] = snapshot.sheets;
  other.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'formula', expression: 'S!A1' } };
  const prefix = 'S!A1&"';
  other.cells['0:1'] = {
    rowId: 0,
    columnId: 1,
    input: { type: 'formula', expression: prefix + 'x'.repeat(32768 - prefix.length - 1) + '"' },
  };
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.sheet.rename',
      payload: { sheetId: sheet.id, name: 'LongerName' },
    }),
  ).toThrowError(expect.objectContaining({ code: 'LIMIT_EXCEEDED' }));
  expect(snapshot).toEqual(before);
});

it('does not consume axis IDs if a merge prevents insertion', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'S', rows: 3, columns: 2 }] });
  const sheet = snapshot.sheets[0];
  sheet.merges.push({ startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.axis.insert',
      payload: { sheetId: sheet.id, axis: 'row', index: 1, ids: [sheet.nextRowId] },
    }),
  ).toThrow();
  expect(snapshot).toEqual(before);
});

it('rejects an oversized sorted formula before changing any cells', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'S', rows: 2, columns: 2 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 2 } };
  sheet.cells['1:0'] = { rowId: 1, columnId: 0, input: { type: 'number', value: 1 } };
  sheet.cells['0:1'] = {
    rowId: 0,
    columnId: 1,
    input: { type: 'formula', expression: `A1&"${'x'.repeat(32763)}"` },
  };
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.sheet.sort',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
        column: 0,
        direction: 'asc',
      },
    }),
  ).toThrowError(expect.objectContaining({ code: 'LIMIT_EXCEEDED' }));
  expect(snapshot).toEqual(before);
});
