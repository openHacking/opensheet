import { it, expect } from 'vitest';
import {
  createWorkbook,
  createSheet,
  createSnapshot,
  Workbook,
  key,
  LIMITS,
  type Selection,
} from '@opensheetjs/core';
import { exportRange } from '@opensheetjs/formats';
it('accepts a Selection as a Rect without letting metadata enter the bounds', async () => {
  const b = await createWorkbook(),
    s = b.getSheets()[0];
  const selection: Selection = {
    sheetId: s.id,
    startRow: 0,
    startColumn: 0,
    endRow: 1,
    endColumn: 1,
  };
  await s.range(selection).setValues([[7]]);
  expect(await s.range(selection).getValues()).toEqual([[7]]);
  expect(s.range(selection).bounds).not.toHaveProperty('sheetId');
});
it('copies overlapping ranges atomically and fills relative formulas', async () => {
  const b = await createWorkbook(),
    s = b.getSheets()[0];
  await s.range('A1:A3').setValues([[1], [2], [3]]);
  await s.range('A1:A3').copyTo('A2:A4');
  expect(await s.range('A1:A4').getValues()).toEqual([[1], [1], [2], [3]]);
  await s.range('B1').setFormulas([['A1*2']]);
  await s.range('B1:B4').fillDown();
  expect(await s.range('B1:B4').getValues()).toEqual([[2], [2], [4], [6]]);
  await b.undo();
  expect(await s.range('B2:B4').getValues()).toEqual([[null], [null], [null]]);
});
it('validates commands that add externally supplied sheets', async () => {
  const b = await createWorkbook(),
    bad = createSheet('Bad');
  bad.rowOrder[1] = bad.rowOrder[0];
  await expect(
    async () => await b.execute({ type: 'core.sheet.add', payload: { sheet: bad } }),
  ).rejects.toThrow();
  expect(b.getSheets()).toHaveLength(1);
  const s = b.getSheets()[0];
  await expect(
    async () =>
      await b.execute({
        type: 'core.axis.insert',
        payload: { sheetId: s.id, axis: 'row', index: 0, ids: ['__proto__'] },
      }),
  ).rejects.toThrow();
  await expect(async () => await b.execute({ type: 'constructor', payload: {} })).rejects.toThrow(
    /Unknown command/,
  );
});
it('exports entire merges and rejects partial merged selections', async () => {
  const b = await createWorkbook(),
    s = b.getSheets()[0];
  await s.range('A1').setValues([['Heading']]);
  await s.range('A1:C3').merge();
  expect(s.getUsedRange().address).toBe('A1:C3');
  await expect(async () =>
    exportRange(await b.toJSON(), { sheetId: s.id, range: 'B1:C3', format: 'html' }),
  ).rejects.toThrow(/entire merged/);
});
it('stores more than the former 200000-cell cap and preserves edits through undo', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Large', rows: 3000, columns: 100 }] });
  await book.generate(210000);
  const sheet = book.getSheets()[0];
  await sheet.range('A1').setValues([[42]]);
  expect(await sheet.range('A1').getValues()).toEqual([[42]]);
  await book.undo();
  expect(await sheet.range('A1').getValues()).toEqual([[0]]);
  expect(book.stats.cacheBytes).toBeLessThanOrEqual(64 * 1024 * 1024);
  await book.deleteStorage();
  book.dispose();
});
it('restores freeze, row height, hidden rows and merges through undo and redo', async () => {
  const book = await createWorkbook(),
    sheet = book.getSheets()[0];
  await sheet.setRowHeight(0, 52);
  await sheet.setFreeze(1, 1);
  await sheet.setRowHidden(1, true);
  await sheet.range('B2:C3').merge();
  const data = () => book.sheetData(sheet.id);
  expect(data().rows[data().rowOrder[0]].size).toBe(52);
  expect(data().freeze).toEqual({ rows: 1, columns: 1 });
  expect(data().rows[data().rowOrder[1]].hidden).toBe(true);
  expect(data().merges).toHaveLength(1);
  for (let i = 0; i < 4; i++) await book.undo();
  expect(data().merges).toHaveLength(0);
  expect(data().rows[data().rowOrder[1]]?.hidden).toBeUndefined();
  expect(data().freeze).toEqual({ rows: 0, columns: 0 });
  expect(data().rows[data().rowOrder[0]]?.size).toBeUndefined();
  for (let i = 0; i < 4; i++) await book.redo();
  expect(data().rows[data().rowOrder[0]].size).toBe(52);
  expect(data().merges).toHaveLength(1);
});
it('keeps calculation and history consistent after a rejected transaction', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Summary' }] });
  const [data, summary] = book.getSheets();
  await data.range('A1').setValues([[2]]);
  await summary.range('A1').setFormulas([['Data!A1*3']]);
  const revision = book.revision;
  expect(await summary.range('A1').getValues()).toEqual([[6]]);
  await expect(
    async () =>
      await book.transaction({}, async (book) => {
        await book
          .getSheetById(data.id)!
          .range('A1')
          .setValues([[5]]);
        expect(await book.getSheetById(summary.id)!.range('A1').getValues()).toEqual([[15]]);
        // A caught command error must still poison the entire transaction.
        try {
          await book.execute({ type: 'core.unknown', payload: {} });
        } catch {
          /* intentionally caught */
        }
      }),
  ).rejects.toThrow();
  expect(book.revision).toBe(revision);
  expect(await data.range('A1').getValues()).toEqual([[2]]);
  expect(await summary.range('A1').getValues()).toEqual([[6]]);
  await book.transaction(
    {},
    async (book) =>
      await book
        .getSheetById(data.id)!
        .range('A1')
        .setValues([[7]]),
  );
  expect(await summary.range('A1').getValues()).toEqual([[21]]);
  await book.undo();
  expect(await summary.range('A1').getValues()).toEqual([[6]]);
  await book.redo();
  expect(await summary.range('A1').getValues()).toEqual([[21]]);
  book.dispose();
  book.dispose();
});
