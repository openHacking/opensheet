import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it.each(['paste', 'replace'])('%s leaves the snapshot intact on a late error', (kind) => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 2, columns: 3 }] });
  const sheet = snapshot.sheets[0];
  sheet.cells['0:1'] = { rowId: 0, columnId: 1, input: { type: 'string', value: 'a' } };
  sheet.cells['0:2'] = {
    rowId: 0,
    columnId: 2,
    input: { type: 'string', value: kind === 'replace' ? 'a'.repeat(32768) : 'protected' },
  };
  const command =
    kind === 'paste'
      ? {
          type: 'core.cells.paste',
          payload: {
            sheetId: sheet.id,
            cells: [
              { row: 0, column: 0, input: { type: 'string', value: 'changed' } },
              { row: 5, column: 0, input: { type: 'blank' } },
            ],
          },
        }
      : {
          type: 'core.cells.replace',
          payload: { sheetId: sheet.id, query: 'a', replacement: 'aa' },
        };
  const before = structuredClone(snapshot);
  expect(() => reduceCommand(snapshot, command)).toThrow();
  expect(snapshot).toEqual(before);
});

it('rejects a later merged paste destination before creating cells or styles', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 2, columns: 3 }] });
  const sheet = snapshot.sheets[0];
  sheet.merges.push({ startRow: 0, endRow: 1, startColumn: 1, endColumn: 3 });
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.cells.paste',
      payload: {
        sheetId: sheet.id,
        cells: [
          { row: 0, column: 0, style: { bold: true } },
          { row: 0, column: 2, input: { type: 'blank' } },
        ],
      },
    }),
  ).toThrow();
  expect(snapshot).toEqual(before);
});
