import { describe, it, expect } from 'vitest';
import { evaluateFormula, parseFormula, offsetFormula } from '@opensheetjs/formula';
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
    ['COUNT(1,"x",TRUE)', 1],
    ['COUNTA(1,"",TRUE)', 3],
  ])('%s', (expression, expected) => expect(calc(String(expression))).toEqual(expected));
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
});
