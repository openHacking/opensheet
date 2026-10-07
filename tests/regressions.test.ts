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
it('accepts a Selection as a Rect without letting metadata enter the bounds', () => {
  const b = createWorkbook(),
    s = b.getSheets()[0];
  const selection: Selection = {
    sheetId: s.id,
    startRow: 0,
    startColumn: 0,
    endRow: 1,
    endColumn: 1,
  };
  s.range(selection).setValues([[7]]);
  expect(s.range(selection).getValues()).toEqual([[7]]);
  expect(s.range(selection).bounds).not.toHaveProperty('sheetId');
});
it('copies overlapping ranges atomically and fills relative formulas', () => {
  const b = createWorkbook(),
    s = b.getSheets()[0];
  s.range('A1:A3').setValues([[1], [2], [3]]);
  s.range('A1:A3').copyTo('A2:A4');
  expect(s.range('A1:A4').getValues()).toEqual([[1], [1], [2], [3]]);
  s.range('B1').setFormulas([['A1*2']]);
  s.range('B1:B4').fillDown();
  expect(s.range('B1:B4').getValues()).toEqual([[2], [2], [4], [6]]);
  b.undo();
  expect(s.range('B2:B4').getValues()).toEqual([[null], [null], [null]]);
});
it('validates commands that add externally supplied sheets', () => {
  const b = createWorkbook(),
    bad = createSheet('Bad');
  bad.rowOrder[1] = bad.rowOrder[0];
  expect(() => b.execute({ type: 'core.sheet.add', payload: { sheet: bad } })).toThrow();
  expect(b.getSheets()).toHaveLength(1);
  const s = b.getSheets()[0];
  expect(() =>
    b.execute({
      type: 'core.axis.insert',
      payload: { sheetId: s.id, axis: 'row', index: 0, ids: ['__proto__'] },
    }),
  ).toThrow();
  expect(() => b.execute({ type: 'constructor', payload: {} })).toThrow(/Unknown command/);
});
it('exports entire merges and rejects partial merged selections', () => {
  const b = createWorkbook(),
    s = b.getSheets()[0];
  s.range('A1').setValues([['Heading']]);
  s.range('A1:C3').merge();
  expect(s.getUsedRange().address).toBe('A1:C3');
  expect(() => exportRange(b.toJSON(), { sheetId: s.id, range: 'B1:C3', format: 'html' })).toThrow(
    /entire merged/,
  );
});
it('enforces the cell cap after edits, deletes, transactions and undo', () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Limit', rows: 10000, columns: 21 }] });
  const data = snapshot.sheets[0];
  for (let n = 0; n < LIMITS.cells; n++) {
    const rowId = data.rowOrder[Math.floor(n / 20)],
      columnId = data.columnOrder[n % 20];
    data.cells[key(rowId, columnId)] = { rowId, columnId, input: { type: 'number', value: n } };
  }
  const book = new Workbook(snapshot),
    sheet = book.getSheets()[0];
  expect(() => sheet.range('U1').setValues([[1]])).toThrow(/Cell limit/);
  sheet.range('A1').clear({ all: true });
  sheet.range('U1').setValues([[1]]);
  expect(sheet.range('U1').getValues()).toEqual([[1]]);
  book.undo();
  expect(sheet.range('U1').getValues()).toEqual([[null]]);
  book.undo();
  expect(sheet.range('A1').getValues()).toEqual([[0]]);
  expect(() => sheet.range('U1').setValues([[2]])).toThrow(/Cell limit/);
  book.transaction({}, () => {
    sheet.range('A1').clear({ all: true });
    sheet.range('U1').setValues([[3]]);
  });
  expect(sheet.range('U1').getValues()).toEqual([[3]]);
}, 20_000);
it('restores freeze, row height, hidden rows and merges through undo and redo', () => {
  const book = createWorkbook(),
    sheet = book.getSheets()[0];
  sheet.setRowHeight(0, 52);
  sheet.setFreeze(1, 1);
  sheet.setRowHidden(1, true);
  sheet.range('B2:C3').merge();
  const data = () => book.sheetData(sheet.id);
  expect(data().rows[data().rowOrder[0]].size).toBe(52);
  expect(data().freeze).toEqual({ rows: 1, columns: 1 });
  expect(data().rows[data().rowOrder[1]].hidden).toBe(true);
  expect(data().merges).toHaveLength(1);
  for (let i = 0; i < 4; i++) book.undo();
  expect(data().merges).toHaveLength(0);
  expect(data().rows[data().rowOrder[1]]?.hidden).toBeUndefined();
  expect(data().freeze).toEqual({ rows: 0, columns: 0 });
  expect(data().rows[data().rowOrder[0]]?.size).toBeUndefined();
  for (let i = 0; i < 4; i++) book.redo();
  expect(data().rows[data().rowOrder[0]].size).toBe(52);
  expect(data().merges).toHaveLength(1);
});
