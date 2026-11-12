import { describe, expect, it } from 'vitest';
import { number } from '../packages/formula/src/coercion.js';
import { FormulaError } from '../packages/formula/src/errors.js';

describe('numeric text radix guard', () => {
  it('matches the previous guard across whitespace, prefixes and short strings', () => {
    const whitespace = ['', ' ', '\t', '\n', '\u00a0', '\ufeff', '\u2003'];
    const beginnings = ['', '0', '00', '0x', '0X', '0b', '0B', '0o', '0O', '+', '-', '1', 'İ'];
    const endings = ['', '0', '1', 'f', '10', '.5', 'e2', 'x', ' '];
    for (const before of whitespace)
      for (const beginning of beginnings)
        for (const ending of endings)
          for (const after of whitespace) {
            const input = before + beginning + ending + after;
            const expected = Number(input);
            const invalid = /^0[xbo]/i.test(input.trim()) || !Number.isFinite(expected);
            if (invalid) {
              try {
                number(input);
                throw new Error(`Accepted ${JSON.stringify(input)}`);
              } catch (error) {
                expect(error).toBeInstanceOf(FormulaError);
                expect((error as FormulaError).code).toBe('#VALUE!');
              }
            } else expect(number(input)).toBe(expected);
          }
  });
});
