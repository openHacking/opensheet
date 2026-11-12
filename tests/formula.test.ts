import { describe, it, expect } from 'vitest';
import {
  createFormulaOffsetter,
  evaluateFormula,
  parseFormula,
  offsetFormula,
} from '@opensheetjs/formula';
import { createWorkbook } from '@opensheetjs/core';
const calc = (s: string) => evaluateFormula(parseFormula(s), () => null);
describe('formulas', () => {
  it.each([
    ['1+2*3', 7],
    ['(1+2)*3', 9],
    ['2^3^2', 512],
    ['SUM(1,2,3)', 6],
    ['AVERAGE(2,4)', 3],
    ['ROUND(-1.25,1)', -1.3],
    ['IF(TRUE,2,1/0)', 2],
    ['IFERROR(1/0,5)', 5],
    ['"a"&"b"', 'ab'],
    ['50%', 0.5],
    ['AND(TRUE,NOT(FALSE))', true],
    ['"X"="x"', true],
    ['COUNT(1,"x",TRUE)', 2],
    ['COUNTA(1,"",TRUE)', 3],
  ])('%s', (expression, expected) => expect(calc(String(expression))).toEqual(expected));
  it('reuses parsed formulas without changing reference relocation at boundaries', () => {
    const formulas = [
      '$A$1+1',
      'A1+$A1+A$1+$A$1',
      'SUM(A1:B2)',
      "'Other Sheet'!$A1+C$3",
      'IF(A1>0,B2,$C$3)',
      '42',
      '#REF!',
      'SUM($A$1:$B$2)',
    ];
    for (const formula of formulas) {
      const offset = createFormulaOffsetter(formula);
      for (const row of [-100000, -2, 0, 1, 99999])
        for (const column of [-1000, -1, 0, 1, 999])
          expect(offset(row, column), `${formula} at ${row},${column}`).toBe(
            offsetFormula(formula, row, column),
          );
    }
    expect(() => createFormulaOffsetter('A1+')).toThrow();
  });
  it('never executes arbitrary code', () => {
    expect(() => parseFormula('globalThis.alert(1);')).toThrow();
    expect(calc('WEBSERVICE("https://example.com")')).toEqual({ error: '#NAME?' });
  });
  it('propagates errors and caps range work', () => {
    expect(calc('1/0')).toEqual({ error: '#DIV/0!' });
    expect(evaluateFormula(parseFormula('SUM(A1:A99999)'), () => 1, 10)).toEqual({
      error: '#LIMIT!',
    });
  });
  it('orders mixed types consistently and bounds concatenated text', () => {
    const values = [0, 'x', false] as const;
    for (let a = 0; a < values.length; a++)
      for (let b = 0; b < values.length; b++) {
        const expression = {
          kind: 'binary',
          op: '<',
          left: { kind: 'literal', value: values[a] },
          right: { kind: 'literal', value: values[b] },
        } as const;
        expect(evaluateFormula(expression, () => null)).toBe(a < b);
      }
    expect(calc('TRUE&"/"&FALSE')).toBe('TRUE/FALSE');
    expect(
      evaluateFormula(
        {
          kind: 'binary',
          op: '&',
          left: { kind: 'literal', value: 'x'.repeat(32767) },
          right: { kind: 'literal', value: 'y' },
        },
        () => null,
      ),
    ).toEqual({ error: '#VALUE!' });
  });
  it('catches cycles and invalidates formula results after every write', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1:B1').setFormulas([['B1', 'A1']]);
    expect(await s.range('A1').getValues()).toEqual([[{ error: '#CYCLE!' }]]);
    await s.range('B1').setValues([[42]]);
    expect(await s.range('A1').getValues()).toEqual([[42]]);
    await b.undo();
    expect(await s.range('A1').getValues()).toEqual([[{ error: '#CYCLE!' }]]);
  });
  it('copies relative references without changing absolute refs or string literals', () => {
    expect(offsetFormula('A1+$B$2+"A1"', 1, 2)).toContain('C2');
    expect(offsetFormula('A1+$B$2+"A1"', 1, 2)).toContain('$B$2');
    expect(offsetFormula('A1+$B$2+"A1"', 1, 2)).toContain('"A1"');
  });
  it('keeps moved references within the parser address limits', () => {
    expect(offsetFormula('XFD1048576', 1, 0)).toBe('#REF!');
    expect(offsetFormula('XFD1048576', 0, 1)).toBe('#REF!');
    expect(offsetFormula('XFC1048575', 1, 1)).toBe('XFD1048576');
    expect(offsetFormula('$XFD$1048576', 1, 1)).toBe('$XFD$1048576');
    expect(offsetFormula('SUM(XFC1:XFD1)', 0, 1)).toBe('SUM(#REF!)');
    expect(offsetFormula('A1', -1, 0)).toBe('#REF!');
  });
  it('persists parseable reference errors after structural insertion', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setFormulas([['A1048576']]);
      await sheet.insertRows(0);
      expect(await sheet.range('A2').getFormulas()).toEqual([['#REF!']]);
      expect(await sheet.range('A2').getValues()).toEqual([[{ error: '#REF!' }]]);
      await book.undo();
      expect(await sheet.range('A1').getFormulas()).toEqual([['A1048576']]);
      await book.redo();
      expect(await sheet.range('A2').getFormulas()).toEqual([['#REF!']]);
    } finally {
      await book.close();
    }
  });
});
