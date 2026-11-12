import { describe, expect, it } from 'vitest';
import {
  formulaReferenceHints,
  formulaReferenceSpans,
  formulaReferenceSpanAt,
} from '@opensheetjs/formula';

describe('formula reference hints', () => {
  it('finds a late selection while preserving the 64-reference highlighting limit', () => {
    const source = `=${Array.from({ length: 100 }, () => 'A1').join('+')}+B2`;
    expect(formulaReferenceSpans(source)).toHaveLength(64);
    const index = source.lastIndexOf('B2');
    expect(formulaReferenceSpanAt(source, index, index + 2)).toMatchObject({
      start: index,
      end: index + 2,
      startRow: 1,
      startColumn: 1,
    });
    expect(formulaReferenceSpanAt(source, index - 1, index + 2)).toBeUndefined();
  });
  it('locates exact UTF-16 spans including quoted sheets and ranges', () => {
    const source = '=IF("😀A1"="x", \'O\'\'Brien\'!$B$2 : A1, C3)';
    expect(formulaReferenceSpans(source).map(({ start, end }) => source.slice(start, end))).toEqual(
      ["'O''Brien'!$B$2 : A1", 'C3'],
    );
  });
  it('extracts incomplete, absolute, reversed and quoted cross-sheet ranges', () => {
    expect(formulaReferenceHints("=SUM($B$2:A1,'O''Brien Q1'!C3:D4,Sheet2!E5")).toEqual([
      { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
      { sheet: "O'Brien Q1", startRow: 2, startColumn: 2, endRow: 4, endColumn: 4 },
      { sheet: 'Sheet2', startRow: 4, startColumn: 4, endRow: 5, endColumn: 5 },
    ]);
  });
  it('ignores text, errors, function names, incomplete quoted sheets and out-of-bounds cells', () => {
    expect(
      formulaReferenceHints('=CONCAT("A1,""B2",LOG10(4),#REF!,XFE1,A1048577,ABC1_name, C3'),
    ).toEqual([{ startRow: 2, startColumn: 2, endRow: 3, endColumn: 3 }]);
    expect(formulaReferenceHints("='Unfinished A1")).toEqual([]);
    expect(formulaReferenceHints('=1E3+1.2+A1')).toEqual([
      { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
    ]);
  });
  it('bounds draft length and number of references', () => {
    expect(formulaReferenceHints('A1')).toEqual([]);
    expect(formulaReferenceHints('=' + 'A1+'.repeat(11000))).toEqual([]);
    expect(formulaReferenceHints('=' + 'A1+'.repeat(100))).toHaveLength(64);
  });
});
