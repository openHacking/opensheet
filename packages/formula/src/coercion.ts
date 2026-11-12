import { fail } from './errors.js';
import type { FormulaValue } from './types.js';

export function checked(v: FormulaValue): FormulaValue {
  if (v !== null && typeof v === 'object') fail(v.error);
  return v;
}
export function number(v: FormulaValue): number {
  checked(v);
  if (v === null || v === '') return 0;
  if (typeof v === 'boolean') return +v;
  // JavaScript accepts programming-language radix prefixes that spreadsheet
  // numeric text does not; retain decimal and exponent coercion.
  if (typeof v === 'string') {
    const text = v.trimStart();
    if (text[0] === '0') {
      const radix = text.charCodeAt(1) | 32;
      if (radix === 120 || radix === 98 || radix === 111) fail();
    }
  }
  const n = Number(v);
  if (!Number.isFinite(n)) fail();
  return n;
}
export function truth(v: FormulaValue): boolean {
  checked(v);
  if (typeof v === 'string') {
    if (v.toUpperCase() === 'TRUE') return true;
    if (v.toUpperCase() === 'FALSE') return false;
    fail();
  }
  return !!v;
}
