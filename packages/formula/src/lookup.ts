import { checked, number, truth } from './coercion.js';
import { wildcard } from './criteria.js';
import { fail } from './errors.js';
import { area, cell, type Area } from './ranges.js';
import type { AST, FormulaValue } from './types.js';

const vectorLength = (range: Area) =>
  range.rows === 1 ? range.columns : range.columns === 1 ? range.rows : fail();
const vectorCell = (range: Area, index: number) =>
  cell(range, range.rows === 1 ? 0 : index, range.rows === 1 ? index : 0);
function compare(
  a: FormulaValue,
  b: FormulaValue,
  work: { remaining: number },
): number | undefined {
  checked(a);
  checked(b);
  if (a === null) a = 0;
  if (b === null) b = 0;
  if (typeof a !== typeof b) return undefined;
  if (typeof a === 'string' && typeof b === 'string') {
    work.remaining -= a.length + b.length;
    if (work.remaining < 0) fail('#LIMIT!');
    a = a.toLowerCase();
    b = b.toLowerCase();
  }
  return a === b ? 0 : a! < b! ? -1 : 1;
}

/** Scalar lookup profile. Iteration retains only the current best candidate. */
export function* lookup(
  name: string,
  args: AST[],
  work = { remaining: 1_000_000 },
): Generator<AST, FormulaValue, FormulaValue> {
  if (name === 'INDEX') {
    const range = area(args[0]);
    let row = Math.trunc(number(yield args[1])),
      column = args[2] ? Math.trunc(number(yield args[2])) : 1;
    if (!args[2] && range.rows === 1) {
      column = row;
      row = 1;
    }
    if (row < 1 || column < 1) return fail(); // Array/reference-valued INDEX is not in this profile.
    if (row > range.rows || column > range.columns) return fail('#REF!');
    return (yield cell(range, row - 1, column - 1)) ?? 0;
  }
  const target = checked(yield args[0]);
  const range = area(args[1]);
  let resultRange: Area | undefined,
    length: number,
    column = 0,
    mode = 0,
    direction = 1,
    ordered = false;
  if (name === 'VLOOKUP') {
    column = Math.trunc(number(yield args[2])) - 1;
    if (column < 0) return fail();
    if (column >= range.columns) return fail('#REF!');
    ordered = !args[3] || truth(yield args[3]);
    if (ordered) mode = -1;
    length = range.rows;
  } else {
    length = vectorLength(range);
    if (name === 'MATCH') {
      const matchType = args[2] ? number(yield args[2]) : 1;
      if (![0, 1, -1].includes(matchType)) return fail();
      ordered = matchType !== 0;
      mode = -matchType;
    } else {
      resultRange = area(args[2]);
      if (resultRange.rows !== range.rows || resultRange.columns !== range.columns) return fail();
      mode = args[4] ? number(yield args[4]) : 0;
      direction = args[5] ? number(yield args[5]) : 1;
      if (
        ![0, -1, 1, 2].includes(mode) ||
        ![1, -1, 2, -2].includes(direction) ||
        (mode === 2 && Math.abs(direction) === 2)
      )
        return fail();
    }
  }
  if (name === 'XLOOKUP' && Math.abs(direction) === 2) {
    // The caller must supply sorted keys; checking every key would defeat binary search.
    let low = 0,
      high = length;
    const ascending = direction === 2;
    while (low < high) {
      const mid = low + Math.floor((high - low) / 2);
      const comparison = compare(checked(yield vectorCell(range, mid)), target, work);
      if (comparison === undefined) return fail();
      if (ascending ? comparison < 0 : comparison > 0) low = mid + 1;
      else high = mid;
    }
    let found = -1;
    if (low < length) {
      const comparison = compare(checked(yield vectorCell(range, low)), target, work);
      if (comparison === undefined) return fail();
      if (comparison === 0) found = low;
    }
    if (found < 0 && mode !== 0) {
      const candidate = (mode === -1) === ascending ? low - 1 : low;
      if (candidate >= 0 && candidate < length) found = candidate;
    }
    if (found < 0) return args[3] ? yield args[3] : { error: '#N/A' };
    return (yield vectorCell(resultRange!, found)) ?? 0;
  }
  const pattern =
    typeof target === 'string' && !ordered && (name !== 'XLOOKUP' || mode === 2)
      ? wildcard(target, work)
      : undefined;
  let found = -1,
    best: FormulaValue = null,
    previous: FormulaValue = null,
    hasPrevious = false;
  for (let step = 0; step < length; step++) {
    const index = direction === -1 ? length - step - 1 : step;
    const value = checked(
      yield name === 'VLOOKUP' ? cell(range, index, 0) : vectorCell(range, index),
    );
    if (ordered && hasPrevious) {
      const order = compare(previous, value, work);
      if (order === undefined || (mode === -1 ? order > 0 : order < 0)) return fail();
    }
    previous = value;
    hasPrevious = true;
    // Exact wildcard searches do not use an ordered comparison.
    const cmp = pattern ? undefined : compare(value, target, work);
    if (ordered && cmp === undefined) return fail();
    const equal = pattern ? typeof value === 'string' && pattern(value) : cmp === 0;
    if (equal) {
      found = index;
      best = value;
      if (!ordered) break;
      continue;
    }
    if (
      (mode === -1 && cmp !== undefined && cmp < 0) ||
      (mode === 1 && cmp !== undefined && cmp > 0)
    ) {
      const comparison = found < 0 ? 0 : compare(value, best, work)!;
      const better =
        found < 0 ||
        (mode === -1 ? comparison > 0 : comparison < 0) ||
        (ordered && comparison === 0);
      if (better) {
        found = index;
        best = value;
      }
    }
  }
  if (found < 0) return name === 'XLOOKUP' && args[3] ? yield args[3] : { error: '#N/A' };
  if (name === 'MATCH') return found + 1;
  return (
    (yield name === 'VLOOKUP' ? cell(range, found, column) : vectorCell(resultRange!, found)) ?? 0
  );
}
