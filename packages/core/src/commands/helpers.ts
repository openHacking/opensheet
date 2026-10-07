import { parseFormula, printFormula, transformReferences } from '@opensheetjs/formula';
import { contains, intersects } from '../address.js';
import { checkBounds } from '../model.js';
import {
  assert,
  key,
  type CellRecord,
  type CellStyle,
  type Rect,
  type SheetSnapshot,
  type WorkbookSnapshot,
} from '../types.js';
export function sheetFor(s: WorkbookSnapshot, id: string) {
  const sh = s.sheets.find((sh) => sh.id === id);
  assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
  return sh;
}
export function styleId(s: WorkbookSnapshot, style: CellStyle): string {
  const value = JSON.stringify(Object.fromEntries(Object.entries(style).sort()));
  let hash = 2166136261;
  for (const c of value) hash = Math.imul(hash ^ c.charCodeAt(0), 16777619);
  let id = `style_${(hash >>> 0).toString(16)}`;
  while (
    s.styles[id] &&
    JSON.stringify(Object.fromEntries(Object.entries(s.styles[id]).sort())) !== value
  )
    id += '_';
  s.styles[id] = style;
  return id;
}
export function cell(sh: SheetSnapshot, row: number, col: number): CellRecord {
  const rowId = sh.rowOrder[row],
    columnId = sh.columnOrder[col],
    k = key(rowId, columnId);
  return sh.cells[k] ?? (sh.cells[k] = { rowId, columnId, input: { type: 'blank' } });
}
export function iterate(sh: SheetSnapshot, r: Rect, fn: (r: number, c: number) => void) {
  checkBounds(sh, r);
  for (let row = r.startRow; row < r.endRow; row++)
    for (let col = r.startColumn; col < r.endColumn; col++) fn(row, col);
}
export function transformAll(
  s: WorkbookSnapshot,
  fn: Parameters<typeof transformReferences>[1],
  sourceSheet?: string,
) {
  for (const sh of s.sheets)
    for (const c of Object.values(sh.cells))
      if (c.input.type === 'formula') {
        try {
          const ast = parseFormula(c.input.expression);
          c.input.expression = printFormula(
            transformReferences(ast, (r) =>
              fn({ ...r, sheet: r.sheet ?? (sourceSheet ? sh.name : undefined) }),
            ),
          );
        } catch {
          assert(
            false,
            'UNSUPPORTED_FEATURE',
            'A formula cannot safely be rewritten; remove it or work on a converted copy',
          );
        }
        delete c.cached;
      }
}
export function guardMerges(sh: SheetSnapshot, r: Rect) {
  for (const m of sh.merges)
    assert(
      !intersects(m, r) || contains(r, m),
      'INVALID_RANGE',
      'Operation cuts through a merged cell',
    );
}
