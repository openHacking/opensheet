import { expect, it } from 'vitest';
import fc from 'fast-check';
import { createWorkbook } from '@opensheetjs/core';
import { createSnapshot } from '../packages/core/src/model.js';
import { reduceCommand } from '../packages/core/src/commands.js';

it('treats positive and negative zero as distinct expected inputs without partial writes', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
  try {
    const sheet = book.getSheets()[0];
    for (const [current, stale] of [
      [0, -0],
      [-0, 0],
    ]) {
      await sheet.range('A1:B1').setValues([[current, 123]]);
      const data = book.sheetData(sheet.id),
        revision = book.revision;
      const expected = {
        rowId: data.rowOrder[0],
        columnId: data.columnOrder[0],
        input: { type: 'number', value: stale },
      };
      const command = {
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [
            { row: 0, column: 1, input: { type: 'number', value: 999 } },
            { row: 0, column: 0, input: { type: 'number', value: 7 }, expected },
          ],
        },
      };
      await expect(book.execute(command)).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
      expect(book.revision).toBe(revision);
      const values = await sheet.range('A1:B1').getValues();
      expect(Object.is(values[0][0], current)).toBe(true);
      expect(values[0][1]).toBe(123);
      expected.input.value = current;
      await book.execute(command);
      expect(await sheet.range('A1:B1').getValues()).toEqual([[7, 999]]);
      await book.undo();
      expect(Object.is((await sheet.range('A1').getValues())[0][0], current)).toBe(true);
    }
  } finally {
    await book.deleteStorage();
  }
});

it('matches an independent atomic model for repeated positions and mixed preconditions', () => {
  fc.assert(
    fc.property(
      fc.array(
        fc.record({
          row: fc.integer({ min: 0, max: 3 }),
          value: fc.integer({ min: -10, max: 10 }),
          mode: fc.integer({ min: 0, max: 2 }),
        }),
        { minLength: 1, maxLength: 30 },
      ),
      (entries) => {
        const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 4, columns: 1 }] });
        const sheet = snapshot.sheets[0];
        for (let row = 0; row < 4; row++)
          sheet.cells[`${row}:0`] = {
            rowId: row,
            columnId: 0,
            input: { type: 'number', value: row },
          };
        const before = structuredClone(sheet.cells);
        const cells = entries.map(({ row, value, mode }) => ({
          row,
          column: 0,
          input: { type: 'number', value },
          ...(mode === 0
            ? {}
            : {
                expected: {
                  rowId: row,
                  columnId: 0,
                  input: { type: 'number', value: row + (mode === 2 ? 1 : 0) },
                },
              }),
        }));
        const command = { type: 'core.cells.set', payload: { sheetId: sheet.id, cells } };
        if (entries.some((entry) => entry.mode === 2)) {
          expect(() => reduceCommand(snapshot, command)).toThrow();
          expect(sheet.cells).toEqual(before);
          return;
        }
        reduceCommand(snapshot, command);
        const expected = [0, 1, 2, 3];
        for (const entry of entries) expected[entry.row] = entry.value;
        for (let row = 0; row < 4; row++)
          expect(sheet.cells[`${row}:0`].input).toEqual({ type: 'number', value: expected[row] });
      },
    ),
    { seed: 712612, numRuns: 500 },
  );
});

it('checks cell identities against the staged sheet and rejects the whole transaction', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[1, 2]]);
    const data = book.sheetData(sheet.id);
    const expected = { rowId: data.rowOrder[0], columnId: data.columnOrder[0] };
    const revision = book.revision;
    await expect(
      book.transaction({}, async (tx) => {
        await tx.getSheets()[0].insertRows(0);
        await tx.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [{ row: 0, column: 0, input: { type: 'number', value: 9 }, expected }],
          },
        });
      }),
    ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1:B1').getValues()).toEqual([[1, 2]]);
    await expect(
      book.transaction({}, async (tx) => {
        await tx.getSheets()[0].insertColumns(0);
        await tx.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [{ row: 0, column: 0, input: { type: 'number', value: 9 }, expected }],
          },
        });
      }),
    ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1:B1').getValues()).toEqual([[1, 2]]);
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 7 }, expected }],
      },
    });
    expect(await sheet.range('A1:B1').getValues()).toEqual([[7, 2]]);
    await book.undo();
    expect(await sheet.range('A1:B1').getValues()).toEqual([[1, 2]]);
  } finally {
    await book.deleteStorage();
  }
});

it('checks duplicate-coordinate preconditions against the same initial command state', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[1]]);
    const data = book.sheetData(sheet.id);
    const identity = { rowId: data.rowOrder[0], columnId: data.columnOrder[0] };
    const command = (secondExpected: number) => ({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [
          {
            row: 0,
            column: 0,
            input: { type: 'number', value: 2 },
            expected: { ...identity, input: { type: 'number', value: 1 } },
          },
          {
            row: 0,
            column: 0,
            input: { type: 'number', value: 3 },
            expected: { ...identity, input: { type: 'number', value: secondExpected } },
          },
        ],
      },
    });
    const revision = book.revision;
    await expect(book.execute(command(2))).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1').getValues()).toEqual([[1]]);
    await book.execute(command(1));
    expect(await sheet.range('A1').getValues()).toEqual([[3]]);
    await book.undo();
    expect(await sheet.range('A1').getValues()).toEqual([[1]]);
    await book.redo();
    expect(await sheet.range('A1').getValues()).toEqual([[3]]);
    await book.transaction({}, async (tx) => {
      await tx.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [
            {
              row: 0,
              column: 0,
              input: { type: 'number', value: 4 },
              expected: { ...identity, input: { type: 'number', value: 3 } },
            },
          ],
        },
      });
      await tx.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [
            {
              row: 0,
              column: 0,
              input: { type: 'number', value: 5 },
              expected: { ...identity, input: { type: 'number', value: 4 } },
            },
          ],
        },
      });
    });
    expect(await sheet.range('A1').getValues()).toEqual([[5]]);
    await book.undo();
    expect(await sheet.range('A1').getValues()).toEqual([[3]]);
  } finally {
    await book.deleteStorage();
  }
});

it('checks expected cell input without blocking unrelated edits', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 4, columns: 4 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[1, 2]]);
    const data = book.sheetData(sheet.id);
    const expected = {
      rowId: data.rowOrder[0],
      columnId: data.columnOrder[0],
      input: { type: 'number', value: 1 },
    };
    await sheet.range('B1').setValues([[3]]);
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 4 }, expected }],
      },
    });
    const revision = book.revision;
    await expect(
      book.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [{ row: 0, column: 0, input: { type: 'number', value: 5 }, expected }],
        },
      }),
    ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1:B1').getValues()).toEqual([[4, 3]]);
    const missing = { rowId: data.rowOrder[1], columnId: data.columnOrder[0], input: null };
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 1, column: 0, input: { type: 'blank' }, expected: missing }],
      },
    });
    await sheet.range('A2').setStyle({ bold: true });
    expect((await book.getCell(sheet.id, 1, 0))?.input).toEqual({ type: 'blank' });
    const blankRevision = book.revision;
    await expect(
      book.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [
            { row: 0, column: 1, input: { type: 'number', value: 10 } },
            { row: 1, column: 0, input: { type: 'number', value: 20 }, expected: missing },
          ],
        },
      }),
    ).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(book.revision).toBe(blankRevision);
    expect(await sheet.range('B1').getValues()).toEqual([[3]]);
  } finally {
    await book.deleteStorage();
  }
});

it('charges expected text against the command planning budget before reading cells', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 600, columns: 1 }] });
  try {
    const sheet = book.getSheets()[0];
    const data = book.sheetData(sheet.id);
    const text = 'x'.repeat(32_768);
    const command = (count: number) => ({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: Array.from({ length: count }, (_, row) => ({
          row,
          column: 0,
          input: { type: 'number', value: 1 },
          expected: {
            rowId: data.rowOrder[row],
            columnId: data.columnOrder[0],
            input: { type: 'string', value: text },
          },
        })),
      },
    });
    const revision = book.revision;
    await expect(book.execute(command(512))).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    await expect(book.execute(command(300))).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
    expect(book.revision).toBe(revision);
    expect(await sheet.range('A1').getValues()).toEqual([[null]]);
  } finally {
    await book.deleteStorage();
  }
});
