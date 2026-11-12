import { assert } from './types.js';

/** Generate from decimal inputs, rather than accumulating binary rounding error. */
export function numericSeries(start: number, step: number, count: number): number[] {
  const decimal = (value: number) => {
    const [mantissa, exponent = '0'] = value.toString().split('e');
    return {
      coefficient: BigInt(mantissa.replace('.', '')),
      exponent: Number(exponent) - (mantissa.split('.')[1]?.length ?? 0),
    };
  };
  const a = decimal(start),
    b = decimal(step),
    exponent = Math.min(a.exponent, b.exponent);
  const first = a.coefficient * 10n ** BigInt(a.exponent - exponent);
  const delta = b.coefficient * 10n ** BigInt(b.exponent - exponent);
  return Array.from({ length: count }, (_, i) => {
    const value = Number(`${first + BigInt(i) * delta}e${exponent}`);
    assert(Number.isFinite(value), 'INVALID_ARGUMENT', 'Series exceeds the finite numeric range');
    return value || 0;
  });
}
