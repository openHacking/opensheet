import { z } from 'zod';
import {
  inputSchema,
  styleSchema,
  rectSchema,
  sheetSchema,
  assert,
  LIMITS,
  key,
  type WorkbookSnapshot,
  type SheetSnapshot,
  type CellStyle,
  type CellRecord,
  type Command,
  type Rect,
} from './types.js';
import { checkBounds, validateSnapshot } from './model.js';
import { contains, intersects } from './address.js';
import {
  parseFormula,
  printFormula,
  transformReferences,
  offsetFormula,
} from '@opensheetjs/formula';
const base = { sheetId: z.string() };
export const commandSchemas = {
  'core.cells.set': z.object({
    ...base,
    cells: z
      .array(
        z.object({
          row: z.number().int().nonnegative(),
          column: z.number().int().nonnegative(),
          input: inputSchema,
        }),
      )
      .max(LIMITS.rangeCells),
  }),
  'core.cells.style': z.object({ ...base, range: rectSchema, style: styleSchema }),
  'core.cells.clear': z.object({ ...base, range: rectSchema, all: z.boolean().default(false) }),
  'core.cells.merge': z.object({
    ...base,
    range: rectSchema,
    discardCoveredValues: z.boolean().default(false),
  }),
  'core.cells.unmerge': z.object({ ...base, range: rectSchema }),
  'core.cells.copy': z.object({ ...base, range: rectSchema, target: rectSchema }),
  'core.cells.fillDown': z.object({ ...base, range: rectSchema }),
  'core.axis.insert': z.object({
    ...base,
    axis: z.enum(['row', 'column']),
    index: z.number().int().nonnegative(),
    ids: z
      .array(
        z
          .string()
          .max(100)
          .regex(/^[\w-]+$/)
          .refine((value) => !['__proto__', 'constructor', 'prototype'].includes(value)),
      )
      .min(1)
      .max(LIMITS.rows),
  }),
  'core.axis.delete': z.object({
    ...base,
    axis: z.enum(['row', 'column']),
    index: z.number().int().nonnegative(),
    count: z.number().int().positive(),
  }),
  'core.axis.meta': z.object({
    ...base,
    axis: z.enum(['row', 'column']),
    index: z.number().int().nonnegative(),
    size: z.number().min(20).max(2000).optional(),
    hidden: z.boolean().optional(),
  }),
  'core.sheet.add': z.object({ sheet: sheetSchema }),
  'core.sheet.remove': z.object(base),
  'core.sheet.rename': z.object({ ...base, name: sheetSchema.shape.name }),
  'core.sheet.freeze': z.object({
    ...base,
    rows: z.number().int().nonnegative(),
    columns: z.number().int().nonnegative(),
  }),
  'core.sheet.filter': z.object({
    ...base,
    filter: z
      .object({ column: z.number().int().nonnegative(), query: z.string().max(1000) })
      .nullable(),
  }),
  'core.sheet.sort': z.object({
    ...base,
    range: rectSchema,
    column: z.number().int().nonnegative(),
    direction: z.enum(['asc', 'desc']),
  }),
} as const;
export function validateCommand(command: Command): Command {
  assert(Object.hasOwn(commandSchemas, command.type), 'UNSUPPORTED_FEATURE', 'Unknown command');
  const schema = commandSchemas[command.type as keyof typeof commandSchemas];
  assert(schema, 'UNSUPPORTED_FEATURE', `Unknown command: ${command.type}`);
  const result = schema.safeParse(command.payload);
  assert(result.success, 'INVALID_ARGUMENT', `Invalid ${command.type} payload`);
  return { type: command.type, payload: result.data };
}
function sheetFor(s: WorkbookSnapshot, id: string) {
  const sh = s.sheets.find((sh) => sh.id === id);
  assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
  return sh;
}
function styleId(s: WorkbookSnapshot, style: CellStyle): string {
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
function cell(sh: SheetSnapshot, row: number, col: number): CellRecord {
  const rowId = sh.rowOrder[row],
    columnId = sh.columnOrder[col],
    k = key(rowId, columnId);
  return sh.cells[k] ?? (sh.cells[k] = { rowId, columnId, input: { type: 'blank' } });
}
function iterate(sh: SheetSnapshot, r: Rect, fn: (r: number, c: number) => void) {
  checkBounds(sh, r);
  for (let row = r.startRow; row < r.endRow; row++)
    for (let col = r.startColumn; col < r.endColumn; col++) fn(row, col);
}
function transformAll(
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
function guardMerges(sh: SheetSnapshot, r: Rect) {
  for (const m of sh.merges)
    assert(
      !intersects(m, r) || contains(r, m),
      'INVALID_RANGE',
      'Operation cuts through a merged cell',
    );
}
export function reduceCommand(s: WorkbookSnapshot, command: Command): void {
  const p = command.payload as any; // Validated by commandSchemas at the public boundary.
  if (command.type === 'core.sheet.add') {
    assert(s.sheets.length < LIMITS.sheets, 'LIMIT_EXCEEDED', 'Sheet limit reached');
    assert(
      !s.sheets.some(
        (sh) => sh.id === p.sheet.id || sh.name.toLowerCase() === p.sheet.name.toLowerCase(),
      ),
      'INVALID_ARGUMENT',
      'Duplicate sheet',
    );
    validateSnapshot({
      ...s,
      sheets: [...s.sheets, p.sheet],
      sheetOrder: [...s.sheetOrder, p.sheet.id],
    });
    s.sheets.push(p.sheet);
    s.sheetOrder.push(p.sheet.id);
    return;
  }
  const sh = sheetFor(s, p.sheetId);
  switch (command.type) {
    case 'core.cells.copy':
    case 'core.cells.fillDown': {
      const source = p.range as Rect;
      const target = (command.type === 'core.cells.fillDown' ? p.range : p.target) as Rect;
      checkBounds(sh, source);
      checkBounds(sh, target);
      assert(
        !sh.merges.some((m) => intersects(m, source) || intersects(m, target)),
        'INVALID_RANGE',
        'Unmerge before copying or filling',
      );
      const fill = command.type === 'core.cells.fillDown';
      assert(
        fill ||
          (source.endRow - source.startRow === target.endRow - target.startRow &&
            source.endColumn - source.startColumn === target.endColumn - target.startColumn),
        'INVALID_RANGE',
        'Copy ranges must have equal dimensions',
      );
      const original = new Map<string, CellRecord>();
      iterate(sh, source, (r, c) => {
        const value = sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
        if (value) original.set(`${r}:${c}`, JSON.parse(JSON.stringify(value)));
      });
      iterate(sh, target, (r, c) => {
        if (fill && r === source.startRow) return;
        const sr = fill ? source.startRow : source.startRow + r - target.startRow;
        const sc = source.startColumn + c - target.startColumn;
        const stored = original.get(`${sr}:${sc}`);
        const k = key(sh.rowOrder[r], sh.columnOrder[c]);
        if (!stored) {
          delete sh.cells[k];
          return;
        }
        const value = JSON.parse(JSON.stringify(stored)) as CellRecord;
        value.rowId = sh.rowOrder[r];
        value.columnId = sh.columnOrder[c];
        if (value.input.type === 'formula')
          value.input.expression = offsetFormula(value.input.expression, r - sr, c - sc);
        delete value.cached;
        sh.cells[k] = value;
      });
      break;
    }
    case 'core.cells.set':
      for (const x of p.cells) {
        checkBounds(sh, {
          startRow: x.row,
          startColumn: x.column,
          endRow: x.row + 1,
          endColumn: x.column + 1,
        });
        const m = sh.merges.find(
          (m) =>
            x.row >= m.startRow &&
            x.row < m.endRow &&
            x.column >= m.startColumn &&
            x.column < m.endColumn,
        );
        assert(
          !m || (x.row === m.startRow && x.column === m.startColumn),
          'INVALID_RANGE',
          'Cannot write a covered merged cell',
        );
        cell(sh, x.row, x.column).input = x.input;
        delete cell(sh, x.row, x.column).cached;
      }
      break;
    case 'core.cells.style':
      guardMerges(sh, p.range);
      iterate(sh, p.range, (r, c) => {
        const x = cell(sh, r, c);
        x.styleId = styleId(s, { ...(x.styleId ? s.styles[x.styleId] : {}), ...p.style });
        if (p.style.numberFormat !== undefined) x.numberFormat = p.style.numberFormat;
      });
      break;
    case 'core.cells.clear':
      guardMerges(sh, p.range);
      iterate(sh, p.range, (r, c) => {
        const k = key(sh.rowOrder[r], sh.columnOrder[c]);
        if (p.all) delete sh.cells[k];
        else if (sh.cells[k]) {
          sh.cells[k].input = { type: 'blank' };
          delete sh.cells[k].cached;
        }
      });
      break;
    case 'core.cells.merge':
      checkBounds(sh, p.range);
      assert(
        !sh.merges.some((m) => intersects(m, p.range)),
        'INVALID_RANGE',
        'Overlapping merged cells',
      );
      iterate(sh, p.range, (r, c) => {
        if (r === p.range.startRow && c === p.range.startColumn) return;
        const x = sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
        assert(
          !x || x.input.type === 'blank' || p.discardCoveredValues,
          'INVALID_ARGUMENT',
          'Merge would discard values',
        );
        if (x) x.input = { type: 'blank' };
      });
      sh.merges.push(p.range);
      break;
    case 'core.cells.unmerge':
      checkBounds(sh, p.range);
      sh.merges = sh.merges.filter((m) => !intersects(m, p.range));
      break;
    case 'core.axis.insert':
    case 'core.axis.delete': {
      const row = p.axis === 'row',
        order = row ? sh.rowOrder : sh.columnOrder,
        insert = command.type === 'core.axis.insert',
        count = insert ? p.ids.length : p.count;
      assert(
        p.index <= (insert ? order.length : order.length - count) && count > 0,
        'INVALID_RANGE',
        'Invalid structural range',
      );
      assert(
        insert
          ? order.length + count <= (row ? LIMITS.rows : LIMITS.columns)
          : order.length - count >= 1,
        'LIMIT_EXCEEDED',
        'Invalid sheet size',
      );
      if (insert)
        assert(
          new Set([...order, ...p.ids]).size === order.length + count,
          'INVALID_ARGUMENT',
          'Duplicate axis IDs',
        );
      const start = row ? 'startRow' : 'startColumn',
        end = row ? 'endRow' : 'endColumn';
      for (const m of sh.merges)
        assert(
          !(p.index < m[end] && p.index + (!insert ? count : 0) > m[start]),
          'INVALID_RANGE',
          'Unmerge cells before editing this axis',
        );
      transformAll(
        s,
        (r) => {
          if (r.sheet !== sh.name) return r;
          const at = row ? r.row : r.column;
          if (!insert && at >= p.index && at < p.index + count) return null;
          const changed =
            at >= p.index ? at + (insert ? count : at >= p.index + count ? -count : 0) : at;
          return { ...r, [row ? 'row' : 'column']: changed };
        },
        sh.name,
      );
      const deleted = insert ? [] : order.slice(p.index, p.index + count);
      order.splice(p.index, insert ? 0 : count, ...(insert ? p.ids : []));
      for (const k of Object.keys(sh.cells))
        if (deleted.includes(row ? sh.cells[k].rowId : sh.cells[k].columnId)) delete sh.cells[k];
      for (const id of deleted) delete (row ? sh.rows : sh.columns)[id];
      sh.merges.forEach((m) => {
        if (m[start] >= p.index) {
          m[start] += insert ? count : -count;
          m[end] += insert ? count : -count;
        }
      });
      const f = row ? 'rows' : 'columns';
      sh.freeze[f] = Math.min(
        order.length,
        sh.freeze[f] > p.index
          ? Math.max(p.index, sh.freeze[f] + (insert ? count : -count))
          : sh.freeze[f],
      );
      if (!row && sh.filter) {
        if (!insert && sh.filter.column >= p.index && sh.filter.column < p.index + count)
          sh.filter = null;
        else if (sh.filter.column >= p.index) sh.filter.column += insert ? count : -count;
      }
      break;
    }
    case 'core.axis.meta': {
      const order = p.axis === 'row' ? sh.rowOrder : sh.columnOrder;
      assert(p.index < order.length, 'INVALID_RANGE', 'Unknown row/column');
      const meta = p.axis === 'row' ? sh.rows : sh.columns;
      meta[order[p.index]] = {
        ...meta[order[p.index]],
        ...(p.size !== undefined ? { size: p.size } : {}),
        ...(p.hidden !== undefined ? { hidden: p.hidden } : {}),
      };
      break;
    }
    case 'core.sheet.remove':
      assert(
        s.sheets.length > 1 && s.sheets.some((x) => x.id !== sh.id && !x.hidden),
        'INVALID_ARGUMENT',
        'Cannot remove last visible sheet',
      );
      transformAll(s, (r) => (r.sheet === sh.name ? null : r));
      s.sheets = s.sheets.filter((x) => x.id !== sh.id);
      s.sheetOrder = s.sheetOrder.filter((id) => id !== sh.id);
      break;
    case 'core.sheet.rename':
      assert(
        !s.sheets.some((x) => x.id !== sh.id && x.name.toLowerCase() === p.name.toLowerCase()),
        'INVALID_ARGUMENT',
        'Duplicate sheet name',
      );
      transformAll(s, (r) => (r.sheet === sh.name ? { ...r, sheet: p.name } : r));
      sh.name = p.name;
      break;
    case 'core.sheet.freeze':
      assert(
        p.rows <= sh.rowOrder.length && p.columns <= sh.columnOrder.length,
        'INVALID_RANGE',
        'Invalid freeze',
      );
      sh.freeze = { rows: p.rows, columns: p.columns };
      break;
    case 'core.sheet.filter':
      assert(
        !p.filter || p.filter.column < sh.columnOrder.length,
        'INVALID_RANGE',
        'Invalid filter column',
      );
      sh.filter = p.filter;
      break;
    case 'core.sheet.sort': {
      checkBounds(sh, p.range);
      assert(
        p.column >= p.range.startColumn && p.column < p.range.endColumn,
        'INVALID_RANGE',
        'Sort column outside range',
      );
      assert(
        !sh.merges.some((m) => intersects(m, p.range)),
        'INVALID_RANGE',
        'Unmerge before sorting',
      );
      const rows = Array.from(
        { length: p.range.endRow - p.range.startRow },
        (_, i) => p.range.startRow + i,
      );
      const value = (r: number) => {
        const x = sh.cells[key(sh.rowOrder[r], sh.columnOrder[p.column])]?.input;
        assert(
          x?.type !== 'formula',
          'UNSUPPORTED_FEATURE',
          'Sorting on formula results is not supported yet',
        );
        return x && 'value' in x ? x.value : null;
      };
      rows.sort((a, b) => {
        const x = value(a),
          y = value(b);
        const cmp =
          x === y
            ? 0
            : x === null
              ? 1
              : y === null
                ? -1
                : typeof x === 'number' && typeof y === 'number'
                  ? x - y
                  : String(x).localeCompare(String(y), 'en');
        return p.direction === 'asc' ? cmp : -cmp;
      });
      const copied = rows.map((r) =>
        Array.from({ length: p.range.endColumn - p.range.startColumn }, (_, i) => {
          const x = sh.cells[key(sh.rowOrder[r], sh.columnOrder[p.range.startColumn + i])];
          return x ? (JSON.parse(JSON.stringify(x)) as CellRecord) : undefined;
        }),
      );
      copied.forEach((cells, i) =>
        cells.forEach((x, j) => {
          const r = p.range.startRow + i,
            c = p.range.startColumn + j,
            k = key(sh.rowOrder[r], sh.columnOrder[c]);
          if (!x) {
            delete sh.cells[k];
            return;
          }
          if (x.input.type === 'formula')
            x.input.expression = offsetFormula(x.input.expression, r - rows[i], 0);
          x.rowId = sh.rowOrder[r];
          x.columnId = sh.columnOrder[c];
          delete x.cached;
          sh.cells[k] = x;
        }),
      );
      break;
    }
    default:
      assert(false, 'UNSUPPORTED_FEATURE', 'Unknown command');
  }
  assert(
    s.sheets.reduce((n, sh) => n + Object.keys(sh.cells).length, 0) <= LIMITS.cells,
    'LIMIT_EXCEEDED',
    'Cell limit reached',
  );
}
