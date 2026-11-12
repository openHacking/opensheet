import { assert, type CellValue, type Rect } from './types.js';

export interface SortColumn {
  column: number;
  direction: 'asc' | 'desc';
}
export function sortColumns(
  payload: { keys?: SortColumn[]; column?: number; direction?: 'asc' | 'desc' },
  range: Rect,
): SortColumn[] {
  const keys = payload.keys ?? [{ column: payload.column!, direction: payload.direction! }];
  assert(
    keys.length > 0 &&
      keys.length <= 64 &&
      new Set(keys.map((key) => key.column)).size === keys.length,
    'INVALID_ARGUMENT',
    'Use 1–64 distinct sort columns',
  );
  for (const key of keys) {
    assert(
      Number.isInteger(key.column) &&
        key.column >= range.startColumn &&
        key.column < range.endColumn,
      'INVALID_RANGE',
      'Sort column outside range',
    );
    assert(
      key.direction === 'asc' || key.direction === 'desc',
      'INVALID_ARGUMENT',
      'Invalid sort direction',
    );
  }
  return keys;
}
export function compareSortKeys(a: CellValue[], b: CellValue[], keys: SortColumn[]): number {
  for (let i = 0; i < keys.length; i++) {
    const result = compareSortValues(a[i], b[i], keys[i].direction);
    if (result) return result;
  }
  return 0;
}

export function sortKey(value: CellValue): CellValue {
  assert(
    !(value && typeof value === 'object' && value.error === '#LIMIT!'),
    'LIMIT_EXCEEDED',
    'Sort key calculation exceeded its work budget',
  );
  return value;
}

/** Total ordering; blanks remain last in either direction and equal keys remain stable. */
export function compareSortValues(a: CellValue, b: CellValue, direction: 'asc' | 'desc'): number {
  if (a === b) return 0;
  if (a === null) return 1;
  if (b === null) return -1;
  const rank = (v: Exclude<CellValue, null>) =>
    typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3;
  const type = rank(a) - rank(b);
  const x = typeof a === 'object' ? a.error : typeof a === 'string' ? a.toLowerCase() : a;
  const y = typeof b === 'object' ? b.error : typeof b === 'string' ? b.toLowerCase() : b;
  const result = type || (x === y ? 0 : x < y ? -1 : 1);
  return direction === 'asc' ? result : -result;
}
