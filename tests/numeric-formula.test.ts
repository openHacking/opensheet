import { describe, expect, it, vi } from 'vitest';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { directedRound, modulo } from '../packages/formula/src/numeric.js';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';

describe('rounding and conditional boundaries', () => {
  it('rounds decimal ties and extreme digit counts through both evaluators and persistence', async () => {
    const cases: Array<[string, number]> = [
      ['ROUND(2.15,1)', 2.2],
      ['ROUND(2.149,1)', 2.1],
      ['ROUND(-1.475,2)', -1.48],
      ['ROUND(1.005,2)', 1.01],
      ['ROUND(626.3,-3)', 1000],
      ['ROUND(-50.55,-2)', -100],
      ['ROUND(1,101)', 1],
      ['ROUND(1,-1000)', 0],
    ];
    for (const [formula, expected] of cases)
      expect(evaluateFormula(parseFormula(formula), () => null)).toBe(expected);
    expect(evaluateFormula(parseFormula('ROUND(1.79E308,-308)'), () => null)).toEqual({
      error: '#NUM!',
    });
    expect(evaluateFormula(parseFormula('ROUND(1,1E300)'), () => null)).toBe(1);
    expect(evaluateFormula(parseFormula('ROUND(1,-1E300)'), () => null)).toBe(0);
    let book = await createWorkbook();
    try {
      await book
        .getSheets()[0]
        .range('A1:A8')
        .setFormulas(cases.map(([formula]) => [formula]));
      expect(await book.getSheets()[0].range('A1:A8').getValues()).toEqual(
        cases.map(([, value]) => [value]),
      );
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      expect(await book.getSheets()[0].range('A1:A8').getValues()).toEqual(
        cases.map(([, value]) => [value]),
      );
    } finally {
      await book.close();
    }
  });
  it('invalidates branch dependencies across edits, undo and reopen', async () => {
    let book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:C1').setValues([[false, 1.005, 2.009]]);
      await sheet.range('D1').setFormulas([['IFS(A1,ROUNDUP(B1,2),TRUE,ROUNDDOWN(C1,2))']]);
      expect(await sheet.range('D1').getValues()).toEqual([[2]]);
      await sheet.range('B1').setValues([[3.001]]);
      expect(await sheet.range('D1').getValues()).toEqual([[2]]);
      await sheet.range('A1').setValues([[true]]);
      expect(await sheet.range('D1').getValues()).toEqual([[3.01]]);
      await book.undo();
      expect(await sheet.range('D1').getValues()).toEqual([[2]]);
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      expect(await book.getSheets()[0].range('D1').getValues()).toEqual([[3.01]]);
    } finally {
      await book.close();
    }
  });
  it('rounds extreme exponents without overflowing a scale or allocating unbounded powers', () => {
    expect(directedRound(5e-324, 323, true)).toBe(1e-323);
    expect(directedRound(5e-324, 323, false)).toBe(0);
    expect(directedRound(1e308, 1e300, true)).toBe(1e308);
    expect(directedRound(-1e308, -1e300, false)).toBe(0);
    expect(modulo(1e308, 1e-308)).toBe(0);
    expect(modulo(5e-324, 1e-323)).toBe(5e-324);
  });

  it('satisfies signed integer remainder identities across both divisor signs', () => {
    for (let n = -100; n <= 100; n++)
      for (let d = -12; d <= 12; d++) {
        if (!d) continue;
        const result = modulo(n, d);
        expect(result).toBe(n - d * Math.floor(n / d) || 0);
        expect(Math.abs(result)).toBeLessThan(Math.abs(d));
      }
  });

  it('never reads unselected branches, and validates IFS pairs before reading arguments', async () => {
    for (const expression of ['IFS(TRUE,7,TRUE,B1)', 'IFS(FALSE,B1,TRUE,7)', 'IFNA(7,B1)']) {
      const read = vi.fn(() => ({ error: '#CYCLE!' }));
      expect(evaluateFormula(parseFormula(expression), read)).toBe(7);
      expect(read).not.toHaveBeenCalled();
      const asyncRead = vi.fn(async (_sheet: string, _row: number, column: number) => ({
        rowId: 0,
        columnId: column,
        input: { type: 'formula' as const, expression: column === 0 ? expression : 'B1' },
      }));
      expect(await new AsyncCalculation(asyncRead, () => undefined).value('sheet', 0, 0)).toBe(7);
      expect(asyncRead).toHaveBeenCalledTimes(1);
    }
    const read = vi.fn(() => true);
    expect(evaluateFormula(parseFormula('IFS(A1,7,B1)'), read)).toEqual({ error: '#VALUE!' });
    expect(read).not.toHaveBeenCalled();
  });
});
