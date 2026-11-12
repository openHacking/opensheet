import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it('preflights all covered cells and clears obsolete formula caches when values are discarded', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 2, columns: 3 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } };
  sheet.cells['0:1'] = {
    rowId: 0,
    columnId: 1,
    input: { type: 'blank' },
    cached: { type: 'number', value: 2 },
    note: 'Keep note',
  };
  sheet.cells['0:2'] = {
    rowId: 0,
    columnId: 2,
    input: { type: 'formula', expression: 'A1*2' },
    cached: { type: 'number', value: 2 },
  };
  const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 3 };
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.cells.merge',
      payload: { sheetId: sheet.id, range, discardCoveredValues: false },
    }),
  ).toThrow('Merge would discard values');
  expect(snapshot).toEqual(before);
  reduceCommand(snapshot, {
    type: 'core.cells.merge',
    payload: { sheetId: sheet.id, range, discardCoveredValues: true },
  });
  expect(sheet.cells['0:0'].input).toEqual({ type: 'number', value: 1 });
  expect(sheet.cells['0:1']).toEqual({
    rowId: 0,
    columnId: 1,
    input: { type: 'blank' },
    note: 'Keep note',
  });
  expect(sheet.cells['0:2']).toEqual({ rowId: 0, columnId: 2, input: { type: 'blank' } });
  expect(sheet.merges).toEqual([range]);
});
