import { z } from 'zod';
export const LIMITS = {
  rows: 100000,
  columns: 1000,
  cells: 200000,
  sheets: 20,
  rangeCells: 200000,
  text: 32768,
};
const id = z
  .string()
  .min(1)
  .max(100)
  .regex(/^[\w-]+$/)
  .refine((value) => !['__proto__', 'constructor', 'prototype'].includes(value));
export const scalarSchema = z.discriminatedUnion('type', [
  z.object({ type: z.literal('blank') }),
  z.object({ type: z.literal('string'), value: z.string().max(LIMITS.text) }),
  z.object({ type: z.literal('number'), value: z.number().finite() }),
  z.object({ type: z.literal('boolean'), value: z.boolean() }),
  z.object({ type: z.literal('error'), code: z.string().max(50) }),
]);
export const inputSchema = z.union([
  scalarSchema,
  z.object({ type: z.literal('formula'), expression: z.string().min(1).max(LIMITS.text) }),
]);
export const styleSchema = z.object({
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z
    .string()
    .regex(/^#[\da-fA-F]{6}$/)
    .optional(),
  background: z
    .string()
    .regex(/^#[\da-fA-F]{6}$/)
    .optional(),
  fontSize: z.number().min(8).max(72).optional(),
  align: z.enum(['left', 'center', 'right']).optional(),
  wrap: z.boolean().optional(),
  border: z.boolean().optional(),
  numberFormat: z.string().max(200).optional(),
});
export const rectSchema = z.object({
  startRow: z.number().int().min(0),
  startColumn: z.number().int().min(0),
  endRow: z.number().int().positive(),
  endColumn: z.number().int().positive(),
});
export const cellSchema = z.object({
  rowId: id,
  columnId: id,
  input: inputSchema,
  styleId: id.optional(),
  numberFormat: z.string().max(200).optional(),
  link: z
    .object({ target: z.string().max(4096), tooltip: z.string().max(1000).optional() })
    .optional(),
  note: z.string().max(LIMITS.text).optional(),
  cached: scalarSchema.optional(),
});
export const sheetSchema = z.object({
  id,
  name: z
    .string()
    .min(1)
    .max(31)
    .regex(/^[^\\/?*\[\]:]+$/),
  rowOrder: z.array(id).min(1).max(LIMITS.rows),
  columnOrder: z.array(id).min(1).max(LIMITS.columns),
  cells: z.record(z.string(), cellSchema),
  rows: z.record(
    z.string(),
    z.object({ size: z.number().min(0).max(2000).optional(), hidden: z.boolean().optional() }),
  ),
  columns: z.record(
    z.string(),
    z.object({ size: z.number().min(0).max(2000).optional(), hidden: z.boolean().optional() }),
  ),
  merges: z.array(rectSchema).max(10000),
  freeze: z.object({ rows: z.number().int().min(0), columns: z.number().int().min(0) }),
  filter: z.object({ column: z.number().int().min(0), query: z.string().max(1000) }).nullable(),
  hidden: z.boolean(),
});
export const snapshotSchema = z.object({
  schemaVersion: z.literal(1),
  workbookId: id,
  revision: z.number().int().nonnegative().max(Number.MAX_SAFE_INTEGER),
  dateSystem: z.enum(['1900', '1904']),
  sheetOrder: z.array(id).min(1).max(LIMITS.sheets),
  sheets: z.array(sheetSchema).min(1).max(LIMITS.sheets),
  styles: z.record(id, styleSchema),
  extensions: z.record(z.string(), z.json()),
});
export type Scalar = z.infer<typeof scalarSchema>;
export type CellInput = z.infer<typeof inputSchema>;
export type CellStyle = z.infer<typeof styleSchema>;
export type CellRecord = z.infer<typeof cellSchema>;
export type SheetSnapshot = z.infer<typeof sheetSchema>;
export type WorkbookSnapshot = z.infer<typeof snapshotSchema>;
export type Rect = z.infer<typeof rectSchema>;
export type Primitive = string | number | boolean | null;
export type CellValue = Primitive | { error: string };
export type Selection = Rect & { sheetId: string };
export type Command = { type: string; payload: Record<string, unknown> };
export type CommandEnvelope = Command & {
  protocolVersion: 1;
  commandId: string;
  workbookId: string;
  baseRevision: number;
};
export type Commit = {
  commandId: string;
  previousRevision: number;
  revision: number;
  source: string;
  commands: Command[];
  changedRanges: Selection[];
};
export class OpenSheetError extends Error {
  constructor(
    public code: string,
    message: string,
    public details?: unknown,
  ) {
    super(message);
    this.name = 'OpenSheetError';
  }
}
export function assert(condition: unknown, code: string, message: string): asserts condition {
  if (!condition) throw new OpenSheetError(code, message);
}
export const key = (rowId: string, columnId: string) => `${rowId}:${columnId}`;
export const uid = (prefix = 'id') =>
  `${prefix}_${globalThis.crypto.randomUUID().replaceAll('-', '')}`;
export function toInput(v: Primitive): Scalar {
  if (v === null) return { type: 'blank' };
  if (typeof v === 'string') return { type: 'string', value: v };
  if (typeof v === 'number') {
    assert(Number.isFinite(v), 'INVALID_ARGUMENT', 'Numbers must be finite');
    return { type: 'number', value: v };
  }
  assert(typeof v === 'boolean', 'INVALID_ARGUMENT', 'Expected a scalar value');
  return { type: 'boolean', value: v };
}
export function scalarValue(input: Scalar): CellValue {
  return input.type === 'blank'
    ? null
    : input.type === 'error'
      ? { error: input.code }
      : input.value;
}
