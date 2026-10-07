import { z } from 'zod';
import { createSnapshot, validateSnapshot, validateJSONValue } from './model.js';
import { BLOCK_ROWS, BLOCK_COLUMNS } from './block-layout.js';
import {
  assert,
  cellSchema,
  sheetSchema,
  styleSchema,
  snapshotSchema,
  scalarSchema,
  key,
  LIMITS,
  type CellRecord,
  type CellInput,
  type WorkbookSnapshot,
} from './types.js';

const index = z.number().int().min(0).max(0xffffffff);
const position = z
  .number()
  .int()
  .min(0)
  .max(BLOCK_ROWS * BLOCK_COLUMNS - 1);
const column = <T extends z.ZodType>(value: T) =>
  z.object({ positions: z.array(position).max(2048), values: z.array(value).max(2048) }).strict();
const axisSchema = z
  .object({
    nextId: z.number().int().min(1).max(0x100000000),
    order: z
      .array(
        z.tuple([
          index,
          z.number().int().positive().max(LIMITS.rows),
          z.union([z.literal(1), z.literal(-1)]),
        ]),
      )
      .max(LIMITS.rows),
    meta: z.record(z.string(), sheetSchema.shape.rows.valueType).optional(),
  })
  .strict();
export const fileBlockDataSchema = z
  .object({
    numbers: column(z.number().finite()).optional(),
    negativeZero: z.array(position).max(2048).optional(),
    strings: column(z.string().max(LIMITS.text)).optional(),
    booleans: column(z.boolean()).optional(),
    formulas: column(z.string().min(1).max(LIMITS.text)).optional(),
    errors: column(z.string().max(50)).optional(),
    blanks: z.array(position).max(2048).optional(),
    styles: z
      .array(z.tuple([position, z.number().int().positive().max(2048), index]))
      .max(2048)
      .optional(),
    notes: column(cellSchema.shape.note.unwrap()).optional(),
    links: column(cellSchema.shape.link.unwrap()).optional(),
    formats: column(cellSchema.shape.numberFormat.unwrap()).optional(),
    cached: column(scalarSchema).optional(),
    cachedNegativeZero: z.array(position).max(2048).optional(),
  })
  .strict();
export const fileBlockSchema = z
  .object({ row: index, column: index, data: fileBlockDataSchema })
  .strict();
const fileSheetSchema = sheetSchema
  .omit({
    rowOrder: true,
    columnOrder: true,
    nextRowId: true,
    nextColumnId: true,
    rows: true,
    columns: true,
    cells: true,
  })
  .extend({
    rows: axisSchema,
    columns: axisSchema,
    blocks: z.array(fileBlockSchema),
  })
  .strict();
export const workbookFileSchema = snapshotSchema
  .omit({ sheets: true, styles: true, sheetOrder: true })
  .extend({
    schemaVersion: z.literal(3),
    sheets: z.array(fileSheetSchema).min(1).max(LIMITS.sheets),
    styles: z.array(styleSchema),
  })
  .strict();
export type WorkbookFile = z.infer<typeof workbookFileSchema>;
export type FileBlock = z.infer<typeof fileBlockSchema>;
export type FileBlockData = z.infer<typeof fileBlockDataSchema>;
export type FileAxis = z.infer<typeof axisSchema>;

export function encodeAxis(
  ids: readonly number[],
  nextId: number,
  meta: Record<string, { size?: number; hidden?: boolean }> = {},
): FileAxis {
  const order: FileAxis['order'] = [];
  for (let i = 0; i < ids.length; ) {
    const step = ids[i + 1] === ids[i] - 1 ? -1 : 1;
    let end = i + 1;
    while (end < ids.length && ids[end] === ids[i] + step * (end - i)) end++;
    order.push([ids[i], end - i, step]);
    i = end;
  }
  return { nextId, order, ...(Object.keys(meta).length ? { meta } : {}) };
}
export function decodeAxis(axis: FileAxis, limit: number): number[] {
  const ids: number[] = [],
    seen = new Set<number>();
  for (const [start, count, step] of axis.order) {
    assert(ids.length + count <= limit, 'LIMIT_EXCEEDED', 'Axis exceeds dimension limit');
    for (let i = 0; i < count; i++) {
      const id = start + i * step;
      assert(
        id >= 0 && id < axis.nextId && !seen.has(id),
        'INVALID_ARGUMENT',
        'Invalid or duplicate axis identity',
      );
      seen.add(id);
      ids.push(id);
    }
  }
  assert(ids.length > 0, 'INVALID_ARGUMENT', 'Empty axis');
  for (const id of Object.keys(axis.meta ?? {}))
    assert(/^(0|[1-9]\d*)$/.test(id) && seen.has(+id), 'INVALID_ARGUMENT', 'Unknown axis metadata');
  return ids;
}
const inputColumns = {
  number: 'numbers',
  string: 'strings',
  boolean: 'booleans',
  formula: 'formulas',
  error: 'errors',
} as const;
/** A single block conversion is shared by object and streaming exports. */
export function encodeFileCells(
  cells: Iterable<CellRecord>,
  styles: Map<string, number>,
  keepCached = true,
): FileBlockData {
  const data: FileBlockData = {};
  const append = (name: string, offset: number, value: unknown) => {
    const col = ((data as any)[name] ??= { positions: [], values: [] });
    col.positions.push(offset);
    col.values.push(value);
  };
  let run: [number, number, number] | undefined;
  for (const cell of [...cells].sort(
    (a, b) => (a.rowId % 64) * 32 + (a.columnId % 32) - (b.rowId % 64) * 32 - (b.columnId % 32),
  )) {
    const offset = (cell.rowId % BLOCK_ROWS) * BLOCK_COLUMNS + (cell.columnId % BLOCK_COLUMNS);
    const input = cell.input;
    if (input.type === 'blank') (data.blanks ??= []).push(offset);
    else
      append(
        inputColumns[input.type],
        offset,
        'value' in input
          ? input.type === 'number' && Object.is(input.value, -0)
            ? 0
            : input.value
          : input.type === 'formula'
            ? input.expression
            : input.code,
      );
    if (input.type === 'number' && Object.is(input.value, -0))
      (data.negativeZero ??= []).push(offset);
    if (cell.styleId !== undefined) {
      const style = styles.get(cell.styleId);
      assert(style !== undefined, 'INVALID_ARGUMENT', 'Unknown style');
      if (run && run[0] + run[1] === offset && run[2] === style) run[1]++;
      else {
        run = [offset, 1, style];
        (data.styles ??= []).push(run);
      }
    } else run = undefined;
    if (cell.note !== undefined) append('notes', offset, cell.note);
    if (cell.link !== undefined) append('links', offset, cell.link);
    if (cell.numberFormat !== undefined) append('formats', offset, cell.numberFormat);
    if (keepCached && cell.cached !== undefined) {
      append(
        'cached',
        offset,
        cell.cached.type === 'number' && Object.is(cell.cached.value, -0)
          ? { ...cell.cached, value: 0 }
          : cell.cached,
      );
      if (cell.cached.type === 'number' && Object.is(cell.cached.value, -0))
        (data.cachedNegativeZero ??= []).push(offset);
    }
  }
  return data;
}
export function decodeFileCells(
  block: FileBlock,
  styles: readonly string[],
  rows?: Set<number>,
  columns?: Set<number>,
): Record<string, CellRecord> {
  let textBytes = 0;
  for (const col of Object.values(block?.data ?? {})) {
    if (col && typeof col === 'object' && 'values' in col && Array.isArray(col.values)) {
      for (const value of col.values) {
        if (typeof value === 'string') textBytes += value.length * 3;
        else if (value && typeof value === 'object') textBytes += JSON.stringify(value).length * 3;
      }
    }
  }
  assert(
    textBytes <= 32 * 1024 * 1024,
    'LIMIT_EXCEEDED',
    'Page text exceeds 32 MiB working budget',
  );
  const parsed = fileBlockSchema.safeParse(block);
  assert(parsed.success, 'INVALID_ARGUMENT', 'Invalid data block');
  const { row, column, data } = parsed.data;
  assert(
    row <= Math.floor(0xffffffff / 64) && column <= Math.floor(0xffffffff / 32),
    'INVALID_ARGUMENT',
    'Block outside identity space',
  );
  const cells = new Map<number, CellRecord>();
  const create = (offset: number, input: CellInput) => {
    const rowId = row * 64 + Math.floor(offset / 32),
      columnId = column * 32 + (offset % 32);
    assert(
      !cells.has(offset) && (!rows || rows.has(rowId)) && (!columns || columns.has(columnId)),
      'INVALID_ARGUMENT',
      'Duplicate or out-of-axis cell',
    );
    const cell = { rowId, columnId, input };
    cells.set(offset, cell);
    return cell;
  };
  for (const [type, name] of Object.entries(inputColumns)) {
    const col = (data as any)[name];
    if (!col) continue;
    assert(col.positions.length === col.values.length, 'INVALID_ARGUMENT', 'Column lengths differ');
    col.positions.forEach((offset: number, i: number) =>
      create(
        offset,
        type === 'formula'
          ? { type, expression: col.values[i] }
          : type === 'error'
            ? { type, code: col.values[i] }
            : ({ type, value: col.values[i] } as CellInput),
      ),
    );
  }
  for (const offset of data.blanks ?? []) create(offset, { type: 'blank' });
  const zero = (positions: number[] | undefined, cached: boolean) => {
    const seen = new Set<number>();
    for (const offset of positions ?? []) {
      const input = cached ? cells.get(offset)?.cached : cells.get(offset)?.input;
      assert(
        input?.type === 'number' && input.value === 0 && !seen.has(offset),
        'INVALID_ARGUMENT',
        'Invalid negative zero',
      );
      input.value = -0;
      seen.add(offset);
    }
  };
  const styled = new Set<number>();
  for (const [start, count, index] of data.styles ?? []) {
    assert(
      styles[index] !== undefined && start + count <= 2048,
      'INVALID_ARGUMENT',
      'Invalid style range',
    );
    for (let offset = start; offset < start + count; offset++) {
      assert(
        cells.has(offset) && !styled.has(offset),
        'INVALID_ARGUMENT',
        'Overlapping or orphan style',
      );
      cells.get(offset)!.styleId = styles[index];
      styled.add(offset);
    }
  }
  for (const [name, property] of [
    ['notes', 'note'],
    ['links', 'link'],
    ['formats', 'numberFormat'],
    ['cached', 'cached'],
  ] as const) {
    const col = data[name],
      seen = new Set<number>();
    if (!col) continue;
    assert(col.positions.length === col.values.length, 'INVALID_ARGUMENT', 'Column lengths differ');
    col.positions.forEach((offset, i) => {
      assert(
        cells.has(offset) && !seen.has(offset),
        'INVALID_ARGUMENT',
        'Duplicate or orphan property',
      );
      (cells.get(offset)! as any)[property] = col.values[i];
      seen.add(offset);
    });
  }
  zero(data.negativeZero, false);
  zero(data.cachedNegativeZero, true);
  return Object.fromEntries(
    [...cells.values()].map((cell) => [key(cell.rowId, cell.columnId), cell]),
  );
}
export function encodeWorkbookFile(snapshot: WorkbookSnapshot): WorkbookFile {
  const entries = Object.entries(snapshot.styles),
    styleIndices = new Map(entries.map(([id], i) => [id, i]));
  return {
    schemaVersion: 3,
    workbookId: snapshot.workbookId,
    revision: snapshot.revision,
    dateSystem: snapshot.dateSystem,
    extensions: snapshot.extensions,
    styles: entries.map(([, style]) => style),
    sheets: snapshot.sheetOrder
      .map((id) => snapshot.sheets.find((sheet) => sheet.id === id)!)
      .map((sh) => {
        const { rowOrder, columnOrder, nextRowId, nextColumnId, rows, columns, cells, ...meta } =
          sh;
        const grouped = new Map<string, CellRecord[]>();
        for (const cell of Object.values(cells)) {
          const address = `${Math.floor(cell.rowId / 64)}/${Math.floor(cell.columnId / 32)}`;
          if (!grouped.has(address)) grouped.set(address, []);
          grouped.get(address)!.push(cell);
        }
        return {
          ...meta,
          rows: encodeAxis(rowOrder, nextRowId, rows),
          columns: encodeAxis(columnOrder, nextColumnId, columns),
          blocks: [...grouped]
            .map(([address, cells]) => {
              const [row, column] = address.split('/').map(Number);
              return { row, column, data: encodeFileCells(cells, styleIndices) };
            })
            .sort((a, b) => a.row - b.row || a.column - b.column),
        };
      }),
  };
}
/** Only metadata is expanded; import processes each page independently. */
export function decodeFileMetadata(input: unknown): WorkbookSnapshot {
  validateJSONValue(input);
  const parsed = workbookFileSchema.safeParse(input);
  assert(parsed.success, 'INVALID_ARGUMENT', 'Invalid workbook file (schemaVersion 3 required)');
  const file = parsed.data;
  return validateSnapshot({
    ...file,
    sheetOrder: file.sheets.map((sh) => sh.id),
    styles: Object.fromEntries(file.styles.map((style, i) => [`s${i}`, style])),
    sheets: file.sheets.map((sh) => ({
      ...sh,
      rowOrder: decodeAxis(sh.rows, LIMITS.rows),
      columnOrder: decodeAxis(sh.columns, LIMITS.columns),
      nextRowId: sh.rows.nextId,
      nextColumnId: sh.columns.nextId,
      rows: sh.rows.meta ?? {},
      columns: sh.columns.meta ?? {},
      cells: {},
    })),
  });
}
/** Internal adapter helper; public loading does not use this materialization. */
export function decodeWorkbookFile(input: unknown): WorkbookSnapshot {
  const file = input as WorkbookFile;
  const snapshot = decodeFileMetadata({
    ...file,
    sheets: file.sheets?.map((sh) => ({ ...sh, blocks: [] })),
  });
  const styles = Object.keys(snapshot.styles);
  for (let i = 0; i < file.sheets.length; i++) {
    const sh = snapshot.sheets[i],
      rows = new Set(sh.rowOrder),
      columns = new Set(sh.columnOrder),
      seen = new Set<string>();
    for (const block of file.sheets[i].blocks) {
      const address = `${block.row}/${block.column}`;
      assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate block');
      seen.add(address);
      Object.assign(sh.cells, decodeFileCells(block, styles, rows, columns));
    }
  }
  return snapshot;
}

export function createWorkbookFile(
  options: Parameters<typeof import('./model.js').createSnapshot>[0] = {},
): WorkbookFile {
  return encodeWorkbookFile(createSnapshot(options));
}
export function validateWorkbookFile(input: unknown): WorkbookFile {
  assert(
    input && typeof input === 'object' && Array.isArray((input as WorkbookFile).sheets),
    'INVALID_ARGUMENT',
    'Invalid workbook file',
  );
  const file = input as WorkbookFile;
  const metadata = decodeFileMetadata({
    ...file,
    sheets: file.sheets.map((sh) => ({ ...sh, blocks: [] })),
  });
  const styles = Object.keys(metadata.styles);
  for (let i = 0; i < file.sheets.length; i++) {
    const sh = metadata.sheets[i],
      rows = new Set(sh.rowOrder),
      columns = new Set(sh.columnOrder),
      seen = new Set<string>();
    assert(Array.isArray(file.sheets[i].blocks), 'INVALID_ARGUMENT', 'Invalid blocks');
    for (const block of file.sheets[i].blocks) {
      const address = `${block.row}/${block.column}`;
      assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate page');
      seen.add(address);
      decodeFileCells(block, styles, rows, columns);
    }
  }
  return file;
}
