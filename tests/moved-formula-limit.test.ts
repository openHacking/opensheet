import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it.each(['core.cells.copy', 'core.cells.fillDown'])(
  '%s rejects a relocated formula that exceeds the text limit without mutation',
  (type) => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 2, columns: 2 }] });
    const sheet = snapshot.sheets[0];
    sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 42 } };
    sheet.cells['0:1'] = {
      rowId: 0,
      columnId: 1,
      input: { type: 'formula', expression: `A1&"${'x'.repeat(32763)}"` },
    };
    const before = structuredClone(snapshot);
    const payload =
      type === 'core.cells.copy'
        ? {
            sheetId: sheet.id,
            range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
            target: { startRow: 1, endRow: 2, startColumn: 0, endColumn: 2 },
          }
        : { sheetId: sheet.id, range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 } };
    expect(() => reduceCommand(snapshot, { type, payload })).toThrowError(
      expect.objectContaining({ code: 'LIMIT_EXCEEDED' }),
    );
    expect(snapshot).toEqual(before);
  },
);
