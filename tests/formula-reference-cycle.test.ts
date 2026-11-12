import { describe, expect, it } from 'vitest';
import { cycleFormulaReference } from '../packages/renderer/src/formula-references.js';

const cycle = (source: string, start: number, end = start) => {
  const edit = cycleFormulaReference(source, start, end);
  return edit && source.slice(0, edit.start) + edit.text + source.slice(edit.end);
};

describe('reference anchoring cycle', () => {
  it('cycles all four anchor modes at the caret', () => {
    expect(cycle('=A1', 3)).toBe('=$A$1');
    expect(cycle('=$A$1', 5)).toBe('=A$1');
    expect(cycle('=A$1', 4)).toBe('=$A1');
    expect(cycle('=$A1', 4)).toBe('=A1');
  });
  it('handles qualified ranges and selected endpoints without changing quoted text', () => {
    const text = `=SUM('A!B''s'!A1:B2,"A1")`;
    const first = text.indexOf('A1:B2');
    expect(cycle(text, first, first + 5)).toBe(`=SUM('A!B''s'!$A$1:$B$2,"A1")`);
    expect(cycle(text, first + 4)).toBe(`=SUM('A!B''s'!A1:$B$2,"A1")`);
    expect(cycle(text, text.lastIndexOf('A1') + 1)).toBeUndefined();
    expect(cycle(text, 7)).toBeUndefined();
  });
  it('ignores functions, error literals, non-formulas and invalid positions', () => {
    expect(cycle('=LOG10(1)', 4)).toBeUndefined();
    expect(cycle('=#REF!', 4)).toBeUndefined();
    expect(cycle('A1', 2)).toBeUndefined();
    expect(cycle('=A1', -1)).toBeUndefined();
    expect(cycle('=A1', 1, 4)).toBeUndefined();
  });
  it('does not grow a draft beyond the supported formula text limit', () => {
    const source = '=A1+' + '1'.repeat(32763);
    expect(source.length).toBe(32767);
    expect(cycle(source, 3)).toBeUndefined();
  });
  it('cycles references after the first 64 without treating quoted text as a reference', () => {
    const source = `=SUM(${Array.from({ length: 5000 }, () => 'A1').join(',')},"Z9",B2)`;
    const target = source.lastIndexOf('B2');
    const result = cycle(source, target + 1);
    expect(result).toBe(source.slice(0, target) + '$B$2)');
    expect(cycle(source, source.indexOf('Z9') + 1)).toBeUndefined();
  });
});
