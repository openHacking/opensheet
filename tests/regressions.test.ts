import { it, expect } from 'vitest';
import { createWorkbook, createSheet, type Selection } from '@opensheetjs/core';
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
