import { describe, expect, it } from 'vitest';
import { canInsertFormulaReference } from '@opensheetjs/renderer';

describe('formula reference insertion context', () => {
  it('accepts operand positions without entering quoted text or sheet names', () => {
    for (const draft of ['=', '=SUM(', '=A1+', '=SUM(A1, '])
      expect(canInsertFormulaReference(draft, draft.length)).toBe(true);
    for (const draft of ['=SUM("text', "='Sheet, A1", '=A1', 'text'])
      expect(canInsertFormulaReference(draft, draft.length)).toBe(false);
  });
  it('bounds draft size and rejects invalid cursor offsets', () => {
    const long = '=' + ' '.repeat(32767);
    expect(canInsertFormulaReference(long, long.length)).toBe(false);
    for (const cursor of [-1, 0, 1.5, NaN, 99])
      expect(canInsertFormulaReference('=SUM(', cursor)).toBe(false);
  });
});
