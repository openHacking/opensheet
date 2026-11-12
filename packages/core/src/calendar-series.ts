import { assert } from './types.js';

/** Calendar steps preserve the original day when possible, clamping in shorter months. */
export function calendarSeries(
  start: number,
  step: number,
  count: number,
  unit: 'month' | 'year',
  system: '1900' | '1904',
): number[] {
  const max = system === '1900' ? 2958465 : 2957003;
  assert(
    Number.isSafeInteger(start) &&
      start >= (system === '1900' ? 1 : 0) &&
      start <= max &&
      Number.isSafeInteger(step),
    'INVALID_ARGUMENT',
    'Calendar series requires an in-range date and an integer step',
  );
  const epoch = system === '1900' ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  const special = system === '1900' && start === 60;
  const date = new Date(epoch + (start - (system === '1900' && start > 60 ? 1 : 0)) * 86400000);
  const year = special ? 1900 : date.getUTCFullYear();
  const month = special ? 1 : date.getUTCMonth();
  const day = special ? 29 : date.getUTCDate();
  const first = year * 12 + month;
  const delta = step * (unit === 'year' ? 12 : 1);
  const last = first + delta * (count - 1);
  assert(
    Number.isSafeInteger(delta) &&
      Number.isSafeInteger(last) &&
      Math.min(first, last) >= (system === '1900' ? 1900 : 1904) * 12 &&
      Math.max(first, last) <= 9999 * 12 + 11,
    'INVALID_ARGUMENT',
    'Calendar series exceeds the workbook date range',
  );
  return Array.from({ length: count }, (_, i) => {
    const index = first + delta * i;
    const y = Math.floor(index / 12);
    const m = index % 12;
    const days =
      y === 1900 && m === 1 && system === '1900'
        ? 29
        : new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    const d = Math.min(day, days);
    if (system === '1900' && y === 1900 && m === 1 && d === 29) return 60;
    const serial = (Date.UTC(y, m, d) - epoch) / 86400000;
    return serial + (system === '1900' && serial >= 60 ? 1 : 0);
  });
}
