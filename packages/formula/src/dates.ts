import { number } from './coercion.js';
import { fail } from './errors.js';
import type { FormulaValue } from './types.js';

export type DateSystem = '1900' | '1904';
const DAY = 86400000;
const epoch = (system: DateSystem) =>
  Date.UTC(system === '1904' ? 1904 : 1899, system === '1904' ? 0 : 11, system === '1904' ? 1 : 31);
const maximum = (system: DateSystem) => (system === '1904' ? 2957003 : 2958465);
function valid(serial: number, system: DateSystem) {
  if (!Number.isFinite(serial) || serial < 0 || serial >= maximum(system) + 1) fail('#NUM!');
  return serial;
}
function fromUTC(time: number, system: DateSystem) {
  const serial = (time - epoch(system)) / DAY;
  return serial + (system === '1900' && serial >= 60 ? 1 : 0);
}
function parts(serial: number, system: DateSystem): [number, number, number] {
  const day = Math.floor(valid(serial, system));
  if (system === '1900' && day === 0) return [1900, 1, 0];
  if (system === '1900' && day === 60) return [1900, 2, 29];
  const date = new Date(epoch(system) + (day - (system === '1900' && day > 60 ? 1 : 0)) * DAY);
  return [date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate()];
}
function serial(value: FormulaValue, system: DateSystem): number {
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [year, month, day] = value.split('-').map(Number);
    if (system === '1900' && value === '1900-02-29') return 60;
    const date = new Date(0);
    date.setUTCFullYear(year, month - 1, day);
    date.setUTCHours(0, 0, 0, 0);
    if (
      date.getUTCFullYear() !== year ||
      date.getUTCMonth() + 1 !== month ||
      date.getUTCDate() !== day
    )
      fail();
    return valid(fromUTC(date.getTime(), system), system);
  }
  return valid(number(value), system);
}

export function evaluateDate(
  name: string,
  values: FormulaValue[],
  system: DateSystem,
): FormulaValue | undefined {
  if (name === 'DATE') {
    let year = Math.trunc(number(values[0]));
    const month = Math.trunc(number(values[1]));
    const day = Math.trunc(number(values[2]));
    if (year < 0 || year > 9999) fail('#NUM!');
    if (year < 1900) year += 1900;
    const first = new Date(0);
    first.setUTCFullYear(year, month - 1, 1);
    first.setUTCHours(0, 0, 0, 0);
    return valid(fromUTC(first.getTime(), system) + day - 1, system);
  }
  if (name === 'DAYS') return serial(values[0], system) - serial(values[1], system);
  const index = ['YEAR', 'MONTH', 'DAY'].indexOf(name);
  if (index < 0) return undefined;
  return parts(serial(values[0], system), system)[index];
}
