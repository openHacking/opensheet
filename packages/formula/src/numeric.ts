import { fail } from './errors.js';

function decimal(value: number) {
  const [mantissa, exponent = '0'] = value.toString().split('e');
  return {
    coefficient: BigInt(mantissa.replace('.', '')),
    exponent: Number(exponent) - (mantissa.split('.')[1]?.length ?? 0),
  };
}

/** Round the shortest round-trippable decimal to nearest, with ties away from zero. */
export function nearestRound(value: number, digits: number): number {
  if (!value || digits >= 324) return value || 0;
  if (digits < -308) return 0;
  const { coefficient, exponent } = decimal(Math.abs(value));
  const shift = exponent + digits;
  if (shift >= 0) return value;
  const divisor = 10n ** BigInt(-shift);
  let quotient = coefficient / divisor;
  if ((coefficient % divisor) * 2n >= divisor) quotient++;
  const result = Math.sign(value) * Number(`${quotient}e${-digits}`);
  return Number.isFinite(result) ? result || 0 : fail('#NUM!');
}

/** Quantize the shortest round-trippable decimal, avoiding binary scale artifacts. */
export function directedRound(value: number, places: number, away: boolean): number {
  const digits = Math.trunc(places);
  if (!value || digits >= 324) return value || 0;
  if (digits < -308) return away ? fail('#NUM!') : 0;
  const { coefficient, exponent } = decimal(Math.abs(value));
  const shift = exponent + digits;
  if (shift >= 0) return value;
  const divisor = 10n ** BigInt(-shift);
  let quotient = coefficient / divisor;
  if (away && coefficient % divisor !== 0n) quotient++;
  const result = Math.sign(value) * Number(`${quotient}e${-digits}`);
  return Number.isFinite(result) ? result || 0 : fail('#NUM!');
}

export function modulo(value: number, divisor: number): number {
  if (!divisor) return fail('#DIV/0!');
  const a = decimal(value),
    b = decimal(divisor);
  const exponent = Math.min(a.exponent, b.exponent);
  const numerator = a.coefficient * 10n ** BigInt(a.exponent - exponent);
  const denominator = b.coefficient * 10n ** BigInt(b.exponent - exponent);
  let remainder = numerator % denominator;
  if (remainder !== 0n && remainder < 0n !== denominator < 0n) remainder += denominator;
  return Number(`${remainder}e${exponent}`) || 0;
}
