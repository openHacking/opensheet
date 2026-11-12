import type { CellValue } from './types.js';

function normalizeFormat(format: string) {
  return format.includes('\\') ? format.replace(/\\([-:/ ])/g, '$1') : format;
}

const datePattern = '(?:yyyy-mm-dd|m{1,2}/d{1,2}/(?:yyyy|yy))';
const dateOnly = new RegExp(`^${datePattern}$`, 'i');
const dateAndTime = new RegExp(`^(${datePattern}[ T])?(h{1,2}):mm(?::(ss))?$`, 'i');
const numberFormatters = new Map<number, Intl.NumberFormat>();

function numberFormatter(decimals: number, grouping: boolean, percent = false): Intl.NumberFormat {
  const key = decimals * 4 + Number(grouping) * 2 + Number(percent);
  let formatter = numberFormatters.get(key);
  if (!formatter) {
    formatter = new Intl.NumberFormat('en-US', {
      style: percent ? 'percent' : 'decimal',
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
      useGrouping: grouping,
    });
    numberFormatters.set(key, formatter);
  }
  return formatter;
}

function calendarDate(day: number, format: string, system: '1900' | '1904'): string {
  if (day < 0) return '#####';
  let year: number, month: number, date: number;
  if (system === '1900' && day === 0) [year, month, date] = [1900, 1, 0];
  else if (system === '1900' && day === 60) [year, month, date] = [1900, 2, 29];
  else {
    const origin = system === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    const d = new Date(origin + (day - (system === '1900' && day > 60 ? 1 : 0)) * 86400000);
    if (Number.isNaN(d.getTime())) return String(day);
    [year, month, date] = [d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()];
  }
  if (/^yyyy-mm-dd$/i.test(format))
    return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(date).padStart(2, '0')}`;
  const tokens = /^(m{1,2})\/(d{1,2})\/(yyyy|yy)$/i.exec(format)!;
  const mo = tokens[1].length === 2 ? String(month).padStart(2, '0') : String(month);
  const da = tokens[2].length === 2 ? String(date).padStart(2, '0') : String(date);
  const yr =
    tokens[3].length === 2 ? String(year % 100).padStart(2, '0') : String(year).padStart(4, '0');
  return `${mo}/${da}/${yr}`;
}

/** Conservative display profile; other codes are retained but may render approximately. */
export function isSupportedNumberFormat(format: string): boolean {
  const code = normalizeFormat(format);
  return (
    /^general$/i.test(code) ||
    /^\$?(?:0|#,##0)(?:\.0{1,12})?%?$/.test(code) ||
    dateOnly.test(code) ||
    dateAndTime.test(code)
  );
}

export function formatValue(
  value: CellValue,
  format?: string,
  dateSystem: '1900' | '1904' = '1900',
): string {
  if (value === null) return '';
  if (typeof value === 'object') return value.error;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number' || !format || /^general$/i.test(format)) return String(value);
  format = normalizeFormat(format);
  const dateTime = dateAndTime.exec(format);
  if (dateTime) {
    if (value < 0) return '#####';
    const seconds = dateTime[3] !== undefined;
    const unitsPerDay = seconds ? 86400 : 1440;
    const rounded = Math.round(value * unitsPerDay);
    if (!Number.isSafeInteger(rounded)) return String(value);
    const day = Math.floor(rounded / unitsPerDay);
    const remainder = rounded % unitsPerDay;
    const hours = Math.floor(remainder / (seconds ? 3600 : 60));
    const minutes = Math.floor((remainder % (seconds ? 3600 : 60)) / (seconds ? 60 : 1));
    const hourText = dateTime[2].length === 2 ? String(hours).padStart(2, '0') : String(hours);
    const time = `${hourText}:${String(minutes).padStart(2, '0')}${seconds ? `:${String(remainder % 60).padStart(2, '0')}` : ''}`;
    return dateTime[1]
      ? `${calendarDate(day, dateTime[1].slice(0, -1), dateSystem)}${dateTime[1].slice(-1)}${time}`
      : time;
  }
  if (dateOnly.test(format)) return calendarDate(Math.floor(value), format, dateSystem);
  if (/[yd]/i.test(format)) {
    if (dateSystem === '1900' && Math.floor(value) === 60) return '1900-02-29';
    const origin = dateSystem === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    const day = value - (dateSystem === '1900' && value >= 60 ? 1 : 0);
    const date = new Date(origin + day * 86400000);
    return Number.isNaN(date.getTime()) ? String(value) : date.toISOString().slice(0, 10);
  }
  const decimals = Math.min(12, format.match(/\.([0#]+)/)?.[1].length ?? 0);
  const percent = format.includes('%'),
    scaled = percent ? value * 100 : value,
    overflow = percent && !Number.isFinite(scaled);
  let text = numberFormatter(decimals, format.includes(','), overflow).format(
    overflow ? value : scaled,
  );
  if (percent && !overflow) text += '%';
  if (format.includes('$')) text = text.startsWith('-') ? '-$' + text.slice(1) : '$' + text;
  return text;
}
