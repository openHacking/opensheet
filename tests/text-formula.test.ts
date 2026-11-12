import { describe, expect, it } from 'vitest';
import { createWorkbook } from '@opensheetjs/core';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { toSheetJS } from '@opensheetjs/adapter-sheetjs';
import { toXlsxFormula } from '../packages/adapter-sheetjs/src/formula.js';

describe('bounded text functions and file formula names', () => {
  it('caps text output while allowing long empty ranges without an accumulating item array', () => {
    expect(evaluateFormula(parseFormula('TEXTJOIN("",FALSE,A1:A10000)'), () => null)).toBe('');
    expect(evaluateFormula(parseFormula('CONCAT(A1,A2)'), () => 'x'.repeat(16384))).toEqual({
      error: '#VALUE!',
    });
    expect(
      evaluateFormula(parseFormula('IFERROR(CONCAT(A1,A2),"overflow")'), () => 'x'.repeat(16384)),
    ).toBe('overflow');
    expect(evaluateFormula(parseFormula('UPPER(A1)'), () => 'ß'.repeat(16384))).toEqual({
      error: '#VALUE!',
    });
    expect(evaluateFormula(parseFormula('LEN(A1)'), () => '🙂'.repeat(20000))).toBe(20000);
    expect(evaluateFormula(parseFormula('RIGHT(A1,1)'), () => 'x'.repeat(40000) + '🙂')).toBe('🙂');
  });

  it('recalculates bounded joins after edit and undo in the Worker', async () => {
    const book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setValues([['x'.repeat(16384)], ['y'.repeat(16384)]]);
      await sheet
        .range('B1:B3')
        .setFormulas([
          ['CONCAT(A1:A2)'],
          ['TEXTJOIN("",TRUE,A1:A2)'],
          ['IFERROR(TEXTJOIN("",TRUE,A1:A2),"overflow")'],
        ]);
      expect(await sheet.range('B1:B3').getValues()).toEqual([
        [{ error: '#VALUE!' }],
        [{ error: '#VALUE!' }],
        ['overflow'],
      ]);
      await sheet.range('A2').setValues([['y'.repeat(16383)]]);
      expect(
        (await sheet.range('B1:B3').getValues()).map(([value]) => (value as string).length),
      ).toEqual([32767, 32767, 32767]);
      await book.undo();
      expect(await sheet.range('B3').getValues()).toEqual([['overflow']]);
      const exported = toSheetJS(await book.toJSON());
      expect(exported.workbook.Sheets[sheet.name].B1.f).toBe('_xlfn.CONCAT(A1:A2)');
      expect(exported.workbook.Sheets[sheet.name].B2.f).toBe('_xlfn.TEXTJOIN("",TRUE,A1:A2)');
    } finally {
      await book.close();
    }
  });

  it('adds OOXML prefixes only to function tokens, preserving quoted text and names', () => {
    const source =
      'CONCAT("CONCAT(\nTEXTJOIN(",\'CONCAT(\'!A1,_xlfn.TEXTJOIN("-",TRUE,"x"),MY.CONCAT(1))';
    expect(toXlsxFormula(source)).toBe('_xlfn.' + source);
    expect(toXlsxFormula('"say ""CONCAT("""&concat ("a","b")')).toBe(
      '"say ""CONCAT("""&_xlfn.CONCAT ("a","b")',
    );
    expect(evaluateFormula(parseFormula('_xlfn.CONCAT("a","b")'), () => null)).toBe('ab');
    expect(evaluateFormula(parseFormula('_xlfn._xlws.UNSUPPORTED(A1)'), () => null)).toEqual({
      error: '#NAME?',
    });
  });
});
