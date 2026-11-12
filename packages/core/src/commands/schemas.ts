import { z } from 'zod';
import { filterSchema } from '../filter-schema.js';
import {
  assert,
  inputSchema,
  linkSchema,
  LIMITS,
  rectSchema,
  sheetSchema,
  styleSchema,
  type Command,
} from '../types.js';
const base = { sheetId: z.string() };
const annotation = {
  ...base,
  range: rectSchema,
  note: z.string().max(LIMITS.text).nullable(),
  link: linkSchema.nullable(),
};
export const commandSchemas = {
  'core.cells.pasteTile': z
    .object({
      ...base,
      range: rectSchema,
      rows: z.number().int().positive().max(LIMITS.rows),
      columns: z.number().int().positive().max(LIMITS.columns),
      cells: z
        .array(
          z
            .object({
              input: inputSchema.optional(),
              style: styleSchema.strict().optional(),
              note: annotation.note.optional(),
              link: annotation.link.optional(),
              source: z
                .object({
                  row: z
                    .number()
                    .int()
                    .nonnegative()
                    .max(LIMITS.rows - 1),
                  column: z
                    .number()
                    .int()
                    .nonnegative()
                    .max(LIMITS.columns - 1),
                })
                .strict()
                .optional(),
            })
            .strict()
            .refine(
              (cell) =>
                (cell.input !== undefined ||
                  cell.style !== undefined ||
                  cell.note !== undefined ||
                  cell.link !== undefined) &&
                (cell.source === undefined || cell.input?.type === 'formula'),
              'Invalid paste tile cell',
            ),
        )
        .min(1)
        .max(LIMITS.rangeCells),
    })
    .strict(),
  'core.cells.paste': z
    .object({
      ...base,
      cells: z
        .array(
          z
            .object({
              row: z.number().int().nonnegative(),
              column: z.number().int().nonnegative(),
              input: inputSchema.optional(),
              style: styleSchema.strict().optional(),
              note: annotation.note.optional(),
              link: annotation.link.optional(),
            })
            .strict()
            .refine(
              (cell) =>
                cell.input !== undefined ||
                cell.style !== undefined ||
                cell.note !== undefined ||
                cell.link !== undefined,
              'Paste needs input, style, note or link',
            ),
        )
        .min(1)
        .max(LIMITS.rangeCells),
    })
    .strict(),
  'core.cells.annotation': z.union([
    z.object({ ...annotation, link: annotation.link.optional() }).strict(),
    z.object({ ...annotation, note: annotation.note.optional() }).strict(),
  ]),
  'core.cells.replace': z.object({
    ...base,
    query: z.string().min(1).max(LIMITS.text),
    replacement: z.string().max(LIMITS.text),
  }),
  'core.cells.set': z.object({
    ...base,
    cells: z
      .array(
        z.object({
          row: z.number().int().nonnegative(),
          column: z.number().int().nonnegative(),
          input: inputSchema,
          expected: z
            .object({
              rowId: z.number().int().nonnegative().max(0xffffffff),
              columnId: z.number().int().nonnegative().max(0xffffffff),
              input: inputSchema.nullable().optional(),
            })
            .strict()
            .optional(),
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
  'core.cells.copySpecial': z
    .object({
      ...base,
      range: rectSchema,
      target: rectSchema,
      mode: z.enum(['values', 'formulas', 'formats']),
    })
    .strict(),
  'core.cells.fillSeries': z
    .object({
      ...base,
      range: rectSchema,
      start: z.number().finite(),
      step: z.number().finite().default(1),
      axis: z.enum(['row', 'column']).default('row'),
      calendar: z.enum(['month', 'year']).optional(),
    })
    .strict(),
  'core.cells.fillDown': z.object({ ...base, range: rectSchema }),
  'core.axis.insert': z.object({
    ...base,
    axis: z.enum(['row', 'column']),
    index: z.number().int().nonnegative(),
    ids: z.array(z.number().int().min(0).max(0xffffffff)).min(1).max(LIMITS.rows),
  }),
  'core.axis.delete': z.object({
    ...base,
    axis: z.enum(['row', 'column']),
    index: z.number().int().nonnegative(),
    count: z.number().int().positive(),
  }),
  'core.axis.resize': z
    .object({
      ...base,
      axis: z.enum(['row', 'column']),
      index: z.number().int().nonnegative(),
      count: z.number().int().positive(),
      size: z.number().min(20).max(2000),
    })
    .strict(),
  'core.axis.visibility': z
    .object({
      ...base,
      axis: z.enum(['row', 'column']),
      index: z.number().int().nonnegative(),
      count: z.number().int().positive(),
      hidden: z.boolean(),
    })
    .strict(),
  'core.axis.visibilityAll': z
    .object({
      ...base,
      axis: z.enum(['row', 'column']),
      hidden: z.boolean(),
    })
    .strict(),
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
    filter: filterSchema,
  }),
  'core.sheet.sort': z.union([
    z
      .object({
        ...base,
        range: rectSchema,
        column: z.number().int().nonnegative(),
        direction: z.enum(['asc', 'desc']),
      })
      .strict(),
    z
      .object({
        ...base,
        range: rectSchema,
        keys: z
          .array(
            z
              .object({
                column: z.number().int().nonnegative(),
                direction: z.enum(['asc', 'desc']),
              })
              .strict(),
          )
          .min(1)
          .max(64),
      })
      .strict(),
  ]),
} as const;
export function validateCommand(command: Command): Command {
  assert(Object.hasOwn(commandSchemas, command.type), 'UNSUPPORTED_FEATURE', 'Unknown command');
  const schema = commandSchemas[command.type as keyof typeof commandSchemas];
  assert(schema, 'UNSUPPORTED_FEATURE', `Unknown command: ${command.type}`);
  const result = schema.safeParse(command.payload);
  assert(result.success, 'INVALID_ARGUMENT', `Invalid ${command.type} payload`);
  return { type: command.type, payload: result.data };
}
