import { columnName, columnIndex } from '@opensheetjs/formula';
import { assert, LIMITS, type Rect } from './types.js';
export { columnName, columnIndex };
export function parseAddress(address: string): { row: number; column: number } {
  const m = /^\$?([a-z]+)\$?([1-9]\d*)$/i.exec(address.trim());
  assert(m, 'INVALID_RANGE', `Invalid address: ${address}`);
  return { row: Number(m[2]) - 1, column: columnIndex(m[1]) };
}
export function address(row: number, column: number): string {
  return `${columnName(column)}${row + 1}`;
}
export function parseRange(input: string | Rect): Rect {
  if (typeof input !== 'string') {
    validateRect(input);
    return {
      startRow: input.startRow,
      startColumn: input.startColumn,
      endRow: input.endRow,
      endColumn: input.endColumn,
    };
  }
  const parts = input.split(':');
  assert(parts.length <= 2, 'INVALID_RANGE', 'Expected A1 or A1:B2');
  const a = parseAddress(parts[0]),
    b = parseAddress(parts[1] ?? parts[0]);
  const r = {
    startRow: Math.min(a.row, b.row),
    startColumn: Math.min(a.column, b.column),
    endRow: Math.max(a.row, b.row) + 1,
    endColumn: Math.max(a.column, b.column) + 1,
  };
  validateRect(r);
  return r;
}
export function rangeAddress(r: Rect): string {
  return (
    address(r.startRow, r.startColumn) +
    (r.endRow - r.startRow === 1 && r.endColumn - r.startColumn === 1
      ? ''
      : `:${address(r.endRow - 1, r.endColumn - 1)}`)
  );
}
export function validateRect(r: Rect): void {
  assert(
    [r.startRow, r.startColumn, r.endRow, r.endColumn].every(Number.isInteger) &&
      r.startRow >= 0 &&
      r.startColumn >= 0 &&
      r.endRow > r.startRow &&
      r.endColumn > r.startColumn,
    'INVALID_RANGE',
    'Invalid range bounds',
  );
  assert(
    r.endRow <= LIMITS.rows && r.endColumn <= LIMITS.columns,
    'LIMIT_EXCEEDED',
    'Range exceeds resource limits',
  );
}
export const intersects = (a: Rect, b: Rect) =>
  a.startRow < b.endRow &&
  a.endRow > b.startRow &&
  a.startColumn < b.endColumn &&
  a.endColumn > b.startColumn;
export const contains = (a: Rect, b: Rect) =>
  a.startRow <= b.startRow &&
  a.endRow >= b.endRow &&
  a.startColumn <= b.startColumn &&
  a.endColumn >= b.endColumn;
