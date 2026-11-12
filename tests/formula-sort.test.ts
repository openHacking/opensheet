import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook, key } from '@opensheetjs/core';
import { compareSortValues } from '../packages/core/src/sort-values.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';
import { Engine } from '../packages/core/src/engine.js';
import { Database } from '../packages/core/src/storage.js';

describe('revision-bound formula sorting', () => {
  it('sorts mixed-direction keys with calculated secondary values and stable ties across runs', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 600 }] });
    try {
      const sheet = book.getSheets()[0];
      const rows = Array.from({ length: 600 }, (_, i) => [i % 3, i % 7, i]);
      await sheet.range('A1:C600').setValues(rows);
      await sheet.range('D1:D600').setFormulas(rows.map((_, i) => [`B${i + 1}*2`]));
      const ids = [...book.sheetData(sheet.id).rowOrder];
      await sheet.sort('A1:D600', [
        { column: 0, direction: 'asc' },
        { column: 3, direction: 'desc' },
      ]);
      const expected = rows
        .map((row) => [...row, row[1] * 2])
        .sort((a, b) => a[0] - b[0] || b[3] - a[3] || a[2] - b[2]);
      expect(await sheet.range('A1:D600').getValues()).toEqual(expected);
      expect(book.sheetData(sheet.id).rowOrder).toEqual(ids);
      await book.undo();
      expect(await sheet.range('A1:C600').getValues()).toEqual(rows);
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(await book.getSheets()[0].range('A1:D600').getValues()).toEqual(expected);
    } finally {
      await book.close();
    }
  }, 30000);

  it('merges large rows through the bounded fallback without losing wide payloads', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Wide', rows: 17, columns: 512 }] });
    try {
      const sheet = book.getSheets()[0];
      const payload = 'x'.repeat(8192);
      await sheet
        .range({ startRow: 0, endRow: 1, startColumn: 0, endColumn: 201 })
        .setValues([[2, ...Array(200).fill(payload)]]);
      await sheet
        .range({ startRow: 16, endRow: 17, startColumn: 0, endColumn: 201 })
        .setValues([[1, ...Array(200).fill(payload + 'y')]]);
      await sheet.sort({ startRow: 0, endRow: 17, startColumn: 0, endColumn: 512 }, 0);
      expect(await sheet.range('A1:B3').getValues()).toEqual([
        [1, payload + 'y'],
        [2, payload],
        [null, null],
      ]);
      await book.undo();
      expect(await sheet.range('A17:B17').getValues()).toEqual([[1, payload + 'y']]);
    } finally {
      await book.close();
    }
  }, 30000);

  it('buffers output pages across inserted row identities without touching columns outside the range', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 140, columns: 5 }] });
    try {
      const sheet = book.getSheets()[0];
      const original = Array.from({ length: 140 }, (_, i) => [
        i,
        i,
        i % 7,
        (i % 7) * 2,
        `keep-${i}`,
      ]);
      await sheet.range('A1:E140').setValues(original);
      await sheet.range('D1:D140').setFormulas(original.map((_, i) => [`C${i + 1}*2`]));
      await sheet.insertRows(33);
      await sheet.range('A34:E34').setValues([[999, 999, 2, 4, 'inserted']]);
      await sheet.range('D34').setFormulas([['C34*2']]);
      original.splice(33, 0, [999, 999, 2, 4, 'inserted']);
      const ids = [...book.sheetData(sheet.id).rowOrder];
      const middle = original
        .slice(1)
        .map((row) => row.slice(1, 4))
        .sort((a, b) => Number(b[1]) - Number(a[1]) || Number(a[0]) - Number(b[0]));
      await sheet.sort('B2:D141', [
        { column: 2, direction: 'desc' },
        { column: 1, direction: 'asc' },
      ]);
      const expected = original.map((row, i) =>
        i === 0 ? row : [row[0], ...middle[i - 1], row[4]],
      );
      expect(await sheet.range('A1:E141').getValues()).toEqual(expected);
      expect(book.sheetData(sheet.id).rowOrder).toEqual(ids);
      await book.undo();
      expect(await sheet.range('A1:E141').getValues()).toEqual(original);
    } finally {
      await book.close();
    }
  });

  it('reports measured sort phases and detaches progress listeners', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A3').setValues([[3], [1], [2]]);
      const events: { phase: string; elapsedMs?: number }[] = [];
      const stop = book.onProgress((progress) => events.push(progress));
      await sheet.sort('A1:A3', 0);
      expect(events.map((event) => event.phase)).toEqual([
        'sort-runs',
        'sort-merge',
        'sort-output',
      ]);
      expect(
        events.every((event) => Number.isFinite(event.elapsedMs) && event.elapsedMs! >= 0),
      ).toBe(true);
      stop();
      await sheet.sort('A1:A3', 0, 'desc');
      expect(events).toHaveLength(3);
    } finally {
      await book.close();
    }
  });

  it('includes hidden rows while retaining row visibility and rejects merged ranges atomically', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B3').setValues([
        [2, 1],
        [1, 2],
        [1, 1],
      ]);
      await sheet.setRowHidden(1, true);
      await sheet.sort('A1:B3', [
        { column: 0, direction: 'asc' },
        { column: 1, direction: 'asc' },
      ]);
      expect(await sheet.range('A1:B3').getValues()).toEqual([
        [1, 1],
        [1, 2],
        [2, 1],
      ]);
      const rowId = book.sheetData(sheet.id).rowOrder[1];
      expect(book.sheetData(sheet.id).rows[rowId].hidden).toBe(true);
      await sheet.range('B1').setValues([[null]]);
      await sheet.range('A1:B1').merge();
      const revision = book.revision;
      await expect(sheet.sort('A1:B3', [{ column: 0, direction: 'asc' }])).rejects.toMatchObject({
        code: 'INVALID_RANGE',
      });
      expect(book.revision).toBe(revision);
    } finally {
      await book.close();
    }
  });

  it('rejects invalid or ambiguous keys without publishing a revision', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B2').setValues([
        [2, 1],
        [1, 2],
      ]);
      const revision = book.revision;
      for (const payload of [
        { keys: [] },
        {
          keys: [
            { column: 0, direction: 'asc' },
            { column: 0, direction: 'desc' },
          ],
        },
        { keys: [{ column: 2, direction: 'asc' }] },
        { keys: [{ column: 0, direction: 'sideways' }] },
        { keys: [{ column: 0, direction: 'asc' }], column: 0, direction: 'asc' },
      ]) {
        await expect(
          book.execute({
            type: 'core.sheet.sort',
            payload: {
              sheetId: sheet.id,
              range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 2 },
              ...payload,
            },
          }),
        ).rejects.toThrow();
        expect(book.revision).toBe(revision);
        expect(await sheet.range('A1:B2').getValues()).toEqual([
          [2, 1],
          [1, 2],
        ]);
      }
    } finally {
      await book.close();
    }
  });

  it('applies multi-key priority in the materialized reducer', () => {
    const snapshot = createSnapshot();
    const sheet = snapshot.sheets[0];
    for (const [row, values] of [
      [2, 1],
      [1, 1],
      [1, 3],
      [1, 3],
    ].entries())
      for (const [column, value] of [...values, row].entries())
        sheet.cells[key(row, column)] = {
          rowId: row,
          columnId: column,
          input: { type: 'number', value },
        };
    reduceCommand(snapshot, {
      type: 'core.sheet.sort',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 3 },
        keys: [
          { column: 0, direction: 'asc' },
          { column: 1, direction: 'desc' },
        ],
      },
    });
    expect([0, 1, 2, 3].map((row) => sheet.cells[key(row, 2)].input)).toEqual(
      [2, 3, 1, 0].map((value) => ({ type: 'number', value })),
    );
  });

  it('freezes pre-sort keys even when absolute references recalculate differently after movement', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A3').setValues([[3], [1], [2]]);
      await sheet.range('B1:B3').setFormulas([['$A$1'], ['$A$2'], ['$A$3']]);
      await sheet.sort('A1:B3', 1);
      expect(await sheet.range('A1:B3').getValues()).toEqual([
        [1, 2],
        [2, 3],
        [3, 1],
      ]);
    } finally {
      await book.close();
    }
  });
  it('aborts rather than sorting exhausted calculation budgets as ordinary errors', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 200 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet
        .range('B1:B140')
        .setFormulas(Array.from({ length: 140 }, (_, i) => [i === 139 ? '1' : `B${i + 2}`]));
      await sheet.range('A1').setFormulas([['B1']]);
      await sheet.range('A2').setValues([[0]]);
      const revision = book.revision;
      await expect(sheet.sort('A1:A2', 0)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
      expect(book.revision).toBe(revision);
      expect(await sheet.range('A2').getValues()).toEqual([[0]]);
    } finally {
      await book.close();
    }
  });
  it('uses staged cross-sheet inputs and preserves formula movement/history/identities', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Other' }] });
    try {
      const [data, other] = book.getSheets();
      await data.range('A1:B3').setValues([
        ['three', 3],
        ['one', 1],
        ['two', 2],
      ]);
      await other.range('A1').setValues([[1]]);
      await data
        .range('C1:C3')
        .setFormulas([['B1*Other!$A$1'], ['B2*Other!$A$1'], ['B3*Other!$A$1']]);
      const ids = [...book.sheetData(data.id).rowOrder];
      const revision = book.revision;
      await book.transaction({}, async (tx) => {
        await tx
          .getSheets()[1]
          .range('A1')
          .setValues([[-1]]);
        await tx.getSheets()[0].sort('A1:C3', 2);
      });
      const sorted = [
        ['three', 3, -3],
        ['two', 2, -2],
        ['one', 1, -1],
      ];
      expect(await data.range('A1:C3').getValues()).toEqual(sorted);
      expect(book.sheetData(data.id).rowOrder).toEqual(ids);
      await expect(
        book.execute({
          protocolVersion: 1,
          commandId: 'stale-sort',
          workbookId: book.id,
          baseRevision: revision,
          type: 'core.sheet.sort',
          payload: {
            sheetId: data.id,
            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 },
            column: 2,
            direction: 'asc',
          },
        }),
      ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
      await book.undo();
      expect(await data.range('A1:A3').getValues()).toEqual([['three'], ['one'], ['two']]);
      expect(await other.range('A1').getValues()).toEqual([[1]]);
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(await book.getSheets()[0].range('A1:C3').getValues()).toEqual(sorted);
    } finally {
      await book.close();
    }
  });

  it('defines a transitive mixed-type order with blanks always last', () => {
    const values = [
      -2,
      2,
      10,
      '',
      '10',
      'a',
      'A',
      false,
      true,
      { error: '#DIV/0!' },
      { error: '#REF!' },
      null,
    ];
    for (const direction of ['asc', 'desc'] as const) {
      for (const a of values)
        for (const b of values)
          for (const c of values) {
            const ab = compareSortValues(a, b, direction),
              bc = compareSortValues(b, c, direction);
            if (ab <= 0 && bc <= 0)
              expect(compareSortValues(a, c, direction)).toBeLessThanOrEqual(0);
          }
      expect([...values].sort((a, b) => compareSortValues(a, b, direction)).at(-1)).toBeNull();
    }
  });

  it('sorts formula errors and equal keys in the materialized reducer', () => {
    const snapshot = createSnapshot();
    const sheet = snapshot.sheets[0];
    for (const [row, expression] of ['2', '1/0', '1+1', '1'].entries())
      sheet.cells[key(row, 0)] = {
        rowId: row,
        columnId: 0,
        input: { type: 'formula', expression },
      };
    reduceCommand(snapshot, {
      type: 'core.sheet.sort',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 1 },
        column: 0,
        direction: 'asc',
      },
    });
    expect([0, 1, 2, 3].map((row) => sheet.cells[key(row, 0)].input)).toEqual(
      ['1', '2', '(1+1)', '(1/0)'].map((expression) => ({ type: 'formula', expression })),
    );
    expect(sheet.cells[key(4, 0)]).toBeUndefined();
  });

  it.each(['scratch', 'merge', 'records'] as const)(
    'cancels %s staging without publishing rows or retaining scratch records',
    async (phase) => {
      const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 300, columns: 2 }] });
      const sheet = snapshot.sheets[0];
      for (let row = 0; row < 300; row++)
        sheet.cells[key(row, 0)] = {
          rowId: row,
          columnId: 0,
          input: { type: 'formula', expression: String(300 - row) },
        };
      const db = await Database.open(`sort-cancel-${crypto.randomUUID()}`);
      const engine = new Engine(db, snapshot.workbookId);
      await engine.open(snapshot);
      const write = db.write.bind(db);
      const spy = vi.spyOn(db, 'write').mockImplementation(async (...args) => {
        await write(...args);
        if (
          args[1].length &&
          (phase === 'merge'
            ? args[0] === 'scratch' && String(args[1][0][0]).includes('/sort/1/')
            : args[0] === phase)
        )
          engine.cancel();
      });
      try {
        await expect(
          engine.execute([
            {
              type: 'core.sheet.sort',
              payload: {
                sheetId: sheet.id,
                range: { startRow: 0, endRow: 300, startColumn: 0, endColumn: 1 },
                column: 0,
                direction: 'asc',
              },
            },
          ]),
        ).rejects.toMatchObject({ code: 'ABORTED' });
        expect(engine.view().snapshot.revision).toBe(0);
        engine.begin();
        expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({
          type: 'formula',
          expression: '300',
        });
        const scratch = [];
        for await (const records of db.scan(
          'scratch',
          IDBKeyRange.bound(`${snapshot.workbookId}/`, `${snapshot.workbookId}/\uffff`),
        ))
          scratch.push(...records);
        expect(scratch).toEqual([]);
      } finally {
        spy.mockRestore();
        engine.close();
      }
    },
  );
});
