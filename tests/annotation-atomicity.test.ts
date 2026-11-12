import { expect, it } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it('rejects an annotation covering a merged cell without changing earlier cells', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 3, columns: 3 }] });
  const sheet = snapshot.sheets[0];
  sheet.merges.push({ startRow: 1, endRow: 3, startColumn: 1, endColumn: 3 });
  const before = structuredClone(snapshot);
  expect(() =>
    reduceCommand(snapshot, {
      type: 'core.cells.annotation',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 },
        note: 'note',
      },
    }),
  ).toThrow();
  expect(snapshot).toEqual(before);
});

it('accepts exactly the ranges that avoid covered merged cells', () => {
  for (let mr = 0; mr < 3; mr++)
    for (let mc = 0; mc < 3; mc++)
      for (let me = mr + 1; me <= 3; me++)
        for (let ce = mc + 1; ce <= 3; ce++)
          for (let r = 0; r < 3; r++)
            for (let c = 0; c < 3; c++)
              for (let re = r + 1; re <= 3; re++)
                for (let e = c + 1; e <= 3; e++) {
                  const snapshot = createSnapshot({
                    sheets: [{ name: 'Sheet', rows: 3, columns: 3 }],
                  });
                  const sheet = snapshot.sheets[0];
                  sheet.merges.push({
                    startRow: mr,
                    endRow: me,
                    startColumn: mc,
                    endColumn: ce,
                  });
                  const command = {
                    type: 'core.cells.annotation',
                    payload: {
                      sheetId: sheet.id,
                      range: { startRow: r, endRow: re, startColumn: c, endColumn: e },
                      note: 'note',
                    },
                  };
                  let covered = false;
                  for (let row = r; row < re; row++)
                    for (let col = c; col < e; col++)
                      if (
                        row >= mr &&
                        row < me &&
                        col >= mc &&
                        col < ce &&
                        (row !== mr || col !== mc)
                      )
                        covered = true;
                  if (covered) {
                    const before = structuredClone(snapshot);
                    expect(() => reduceCommand(snapshot, command)).toThrow();
                    expect(snapshot).toEqual(before);
                  } else {
                    reduceCommand(snapshot, command);
                    expect(Object.keys(sheet.cells)).toHaveLength((re - r) * (e - c));
                  }
                }
});
