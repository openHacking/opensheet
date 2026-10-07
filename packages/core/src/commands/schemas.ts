import { z } from 'zod';
import {
  assert,
  inputSchema,
  LIMITS,
  rectSchema,
  sheetSchema,
  styleSchema,
  type Command,
} from '../types.js';
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
