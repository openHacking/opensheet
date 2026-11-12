import {
  offsetFormula,
  parseFormula,
  printFormula,
  transformReferences,
} from '@opensheetjs/formula';
import { contains, intersects } from '../address.js';
import { transformDeletedFormulaRange } from '../deleted-formula-range.js';
import { checkBounds } from '../model.js';
import {
  assert,
  LIMITS,
  key,
  type CellRecord,
  type CellInput,
  type CellStyle,
  type Command,
  type Rect,
  type SheetSnapshot,
  type WorkbookSnapshot,
} from '../types.js';
export function offsetCellFormula(expression: string, rows: number, columns: number): string {
  const result = offsetFormula(expression, rows, columns);
  assert(result.length <= LIMITS.text, 'LIMIT_EXCEEDED', 'Moved formula exceeds the text limit');
  return result;
}
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
export type PasteCell = {
  row: number;
  column: number;
  input?: CellInput;
  style?: CellStyle;
  note?: string | null;
  link?: CellRecord['link'] | null;
};
export function applyPastedCell(s: WorkbookSnapshot, target: CellRecord, item: PasteCell): void {
  if (item.input !== undefined) {
    target.input = item.input;
    delete target.cached;
  }
  if (item.style !== undefined) {
    delete target.numberFormat;
    if (Object.keys(item.style).length) target.styleId = styleId(s, item.style);
    else delete target.styleId;
  }
  if (item.note !== undefined) {
    if (item.note === null) delete target.note;
    else target.note = item.note;
  }
  if (item.link !== undefined) {
    if (item.link === null) delete target.link;
    else target.link = { ...item.link };
  }
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
  rangeTransform?: StructuralRangeTransform,
  skip?: (sheet: SheetSnapshot, cell: CellRecord) => boolean,
) {
  const transform = createCellFormulaTransformer(fn, rangeTransform);
  for (const sh of s.sheets)
    transform(
      sh.cells,
      sourceSheet ? sh.name : undefined,
      true,
      skip && ((cell) => skip(sh, cell)),
    );
  for (const sh of s.sheets)
    transform(
      sh.cells,
      sourceSheet ? sh.name : undefined,
      false,
      skip && ((cell) => skip(sh, cell)),
    );
}

type Reference = Extract<ReturnType<typeof parseFormula>, { kind: 'ref' }>;
type StructuralRangeTransform = (
  start: Reference,
  end: Reference,
  localSheet?: string,
) => [Reference, Reference] | null;

export function createCellFormulaTransformer(
  fn: Parameters<typeof transformReferences>[1],
  rangeTransform?: StructuralRangeTransform,
) {
  const bySheet = new Map<string | undefined, Map<string, string>>();
  let remaining = 256 * 1024;
  return (
    cells: Record<string, CellRecord>,
    localSheet?: string,
    validateOnly = false,
    skip?: (cell: CellRecord) => boolean,
  ) => {
    let cache = bySheet.get(localSheet);
    for (const c of Object.values(cells))
      if (c.input.type === 'formula' && !skip?.(c)) {
        const original = c.input.expression;
        let rewritten = cache?.get(original);
        try {
          if (rewritten === undefined) {
            const ast = parseFormula(original);
            rewritten = printFormula(
              transformReferences(
                ast,
                (r) => fn({ ...r, sheet: r.sheet ?? localSheet }),
                rangeTransform ? (start, end) => rangeTransform(start, end, localSheet) : undefined,
              ),
            );
            const cost = 96 + 2 * (original.length + rewritten.length + (localSheet?.length ?? 0));
            if (cost <= remaining) {
              remaining -= cost;
              if (!cache) bySheet.set(localSheet, (cache = new Map()));
              cache.set(original, rewritten);
            }
          }
        } catch {
          assert(
            false,
            'UNSUPPORTED_FEATURE',
            'A formula cannot safely be rewritten; remove it or work on a converted copy',
          );
        }
        assert(
          rewritten.length <= LIMITS.text,
          'LIMIT_EXCEEDED',
          'Rewritten formula exceeds the text limit',
        );
        if (!validateOnly) {
          c.input.expression = rewritten;
          delete c.cached;
        }
      }
  };
}

/** Shrink a range around a deleted span, preserving endpoint order and anchoring. */
export function structuralRangeReferences(
  command: Command,
  sheetName: string,
): StructuralRangeTransform | undefined {
  if (command.type !== 'core.axis.delete') return undefined;
  const payload = command.payload as { axis: 'row' | 'column'; index: number; count: number };
  const coordinate = payload.axis === 'row' ? 'row' : 'column';
  const source = sheetName.toLowerCase();
  const transform = structuralReferences(command, sheetName);
  return (start, end, localSheet) => {
    const startSheet = (start.sheet ?? localSheet)?.toLowerCase();
    const endSheet = (end.sheet ?? localSheet)?.toLowerCase();
    if (startSheet !== source || endSheet !== source) {
      const a = transform({ ...start, sheet: start.sheet ?? localSheet });
      const b = transform({ ...end, sheet: end.sheet ?? localSheet });
      return a && b ? [a, b] : null;
    }
    const range = transformDeletedFormulaRange(
      start,
      end,
      coordinate,
      payload.index,
      payload.count,
    );
    return (
      range && [
        { ...range[0], sheet: range[0].sheet ?? localSheet },
        { ...range[1], sheet: range[1].sheet ?? localSheet },
      ]
    );
  };
}

/** Shared by the snapshot reducer and the paged engine after command validation. */
export function structuralReferences(
  command: Command,
  sheetName: string,
): Parameters<typeof transformReferences>[1] {
  const name = sheetName.toLowerCase();
  const payload = command.payload as any;
  return (ref) => {
    if (ref.sheet?.toLowerCase() !== name) return ref;
    if (command.type === 'core.sheet.remove') return null;
    if (command.type === 'core.sheet.rename') return { ...ref, sheet: payload.name };
    const coordinate = payload.axis === 'row' ? 'row' : 'column';
    const at = ref[coordinate];
    if (at < payload.index) return ref;
    if (command.type === 'core.axis.insert')
      return { ...ref, [coordinate]: at + payload.ids.length };
    return at < payload.index + payload.count ? null : { ...ref, [coordinate]: at - payload.count };
  };
}
export function guardMerges(sh: SheetSnapshot, r: Rect) {
  for (const m of sh.merges)
    assert(
      !intersects(m, r) || contains(r, m),
      'INVALID_RANGE',
      'Operation cuts through a merged cell',
    );
}
