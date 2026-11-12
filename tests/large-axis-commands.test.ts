import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it('inserts and deletes 99,999 rows while preserving IDs, cells and metadata', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 1, columns: 1 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 7 } };
  sheet.rows[0] = { size: 30 };
  const ids = Array.from({ length: 99999 }, (_, i) => i + 1);
  reduceCommand(snapshot, {
    type: 'core.axis.insert',
    payload: { sheetId: sheet.id, axis: 'row', index: 0, ids },
  });
  expect(sheet.rowOrder).toEqual([...ids, 0]);
  expect(sheet.nextRowId).toBe(100000);
  for (const id of ids)
    sheet.cells[`${id}:0`] = { rowId: id, columnId: 0, input: { type: 'number', value: id } };
  sheet.rows[1] = { hidden: true };
  reduceCommand(snapshot, {
    type: 'core.axis.delete',
    payload: { sheetId: sheet.id, axis: 'row', index: 0, count: ids.length },
  });
  expect(sheet.rowOrder).toEqual([0]);
  expect(sheet.nextRowId).toBe(100000);
  expect(sheet.rows).toEqual({ 0: { size: 30 } });
  expect(sheet.cells).toEqual({
    '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 7 } },
  });
});
