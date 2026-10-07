import type { CellValue } from './types.js';

export function formatValue(
  value: CellValue,
  format?: string,
  dateSystem: '1900' | '1904' = '1900',
): string {
  if (value === null) return '';
  if (typeof value === 'object') return value.error;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number' || !format || format === 'General') return String(value);
  if (/[yd]/i.test(format)) {
    if (dateSystem === '1900' && Math.floor(value) === 60) return '1900-02-29';
    const origin = dateSystem === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    const day = value - (dateSystem === '1900' && value >= 60 ? 1 : 0);
    const date = new Date(origin + day * 86400000);
    return Number.isNaN(date.getTime()) ? String(value) : date.toISOString().slice(0, 10);
  }
  const decimals = Math.min(12, format.match(/\.([0#]+)/)?.[1].length ?? 0);
  let text = (format.includes('%') ? value * 100 : value).toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    useGrouping: format.includes(','),
  });
  if (format.includes('%')) text += '%';
  if (format.includes('$')) text = '$' + text;
  return text;
}
