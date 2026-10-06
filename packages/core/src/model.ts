import {
  snapshotSchema,
  assert,
  LIMITS,
  key,
  uid,
  type WorkbookSnapshot,
  type SheetSnapshot,
  type Rect,
} from './types.js';
import { intersects, validateRect } from './address.js';
export function createSheet(name: string, rows = 100, columns = 26): SheetSnapshot {
  assert(
    Number.isInteger(rows) &&
      rows > 0 &&
      rows <= LIMITS.rows &&
      Number.isInteger(columns) &&
      columns > 0 &&
      columns <= LIMITS.columns,
    'LIMIT_EXCEEDED',
    'Invalid sheet dimensions',
  );
  const id = uid('sheet');
  return {
    id,
    name,
    rowOrder: Array.from({ length: rows }, (_, i) => `${id}_r${i}`),
    columnOrder: Array.from({ length: columns }, (_, i) => `${id}_c${i}`),
    cells: {},
    rows: {},
    columns: {},
    merges: [],
    freeze: { rows: 0, columns: 0 },
    filter: null,
    hidden: false,
  };
}
export function createSnapshot(
  options: { sheets?: Array<{ name: string; rows?: number; columns?: number }> } = {},
): WorkbookSnapshot {
  const sheets = (options.sheets ?? [{ name: 'Sheet1' }]).map((s) =>
    createSheet(s.name, s.rows, s.columns),
  );
  return validateSnapshot({
    schemaVersion: 1,
    workbookId: uid('wb'),
    revision: 0,
    dateSystem: '1900',
    sheetOrder: sheets.map((s) => s.id),
    sheets,
    styles: {},
    extensions: {},
  });
}
export function validateSnapshot(input: unknown): WorkbookSnapshot {
  // Validate hostile JSON before recursive schemas or object merges.
  let nodes = 0;
  const walk = (v: unknown, depth: number) => {
    assert(++nodes < 3000000 && depth < 32, 'LIMIT_EXCEEDED', 'Snapshot is too complex');
    if (v && typeof v === 'object') {
      for (const [k, x] of Object.entries(v)) {
        assert(
          !['__proto__', 'constructor', 'prototype'].includes(k),
          'INVALID_ARGUMENT',
          'Unsafe object key',
        );
        walk(x, depth + 1);
      }
    }
  };
  walk(input, 0);
  const parsed = snapshotSchema.safeParse(input);
  assert(parsed.success, 'INVALID_ARGUMENT', 'Invalid workbook snapshot');
  const s = parsed.data;
  assert(
    new Set(s.sheetOrder).size === s.sheets.length &&
      s.sheetOrder.length === s.sheets.length &&
      s.sheetOrder.every((id) => s.sheets.some((sh) => sh.id === id)),
    'INVALID_ARGUMENT',
    'Invalid sheet order',
  );
  assert(
    new Set(s.sheets.map((sh) => sh.id)).size === s.sheets.length &&
      new Set(s.sheets.map((sh) => sh.name.toLowerCase())).size === s.sheets.length,
    'INVALID_ARGUMENT',
    'Duplicate sheet id/name',
  );
  assert(
    s.sheets.some((sh) => !sh.hidden),
    'INVALID_ARGUMENT',
    'At least one sheet must be visible',
  );
  let count = 0;
  for (const sh of s.sheets) {
    const rows = new Set(sh.rowOrder),
      cols = new Set(sh.columnOrder);
    assert(
      rows.size === sh.rowOrder.length && cols.size === sh.columnOrder.length,
      'INVALID_ARGUMENT',
      'Duplicate row/column identity',
    );
    for (const [k, c] of Object.entries(sh.cells)) {
      assert(++count <= LIMITS.cells, 'LIMIT_EXCEEDED', 'Too many cells');
      assert(
        rows.has(c.rowId) && cols.has(c.columnId) && k === key(c.rowId, c.columnId),
        'INVALID_ARGUMENT',
        'Invalid cell identity',
      );
      assert(!c.styleId || !!s.styles[c.styleId], 'INVALID_ARGUMENT', 'Unknown style');
    }
    for (const id of Object.keys(sh.rows))
      assert(rows.has(id), 'INVALID_ARGUMENT', 'Unknown row metadata');
    for (const id of Object.keys(sh.columns))
      assert(cols.has(id), 'INVALID_ARGUMENT', 'Unknown column metadata');
    assert(
      sh.freeze.rows <= sh.rowOrder.length && sh.freeze.columns <= sh.columnOrder.length,
      'INVALID_ARGUMENT',
      'Invalid freeze',
    );
    if (sh.filter)
      assert(sh.filter.column < sh.columnOrder.length, 'INVALID_ARGUMENT', 'Invalid filter');
    sh.merges.forEach((r, i) => {
      checkBounds(sh, r);
      assert(
        !sh.merges.slice(0, i).some((x) => intersects(x, r)),
        'INVALID_ARGUMENT',
        'Overlapping merges',
      );
    });
  }
  return s;
}
export function checkBounds(sheet: SheetSnapshot, rect: Rect): void {
  validateRect(rect);
  assert(
    rect.endRow <= sheet.rowOrder.length && rect.endColumn <= sheet.columnOrder.length,
    'INVALID_RANGE',
    'Range exceeds sheet dimensions',
  );
}
