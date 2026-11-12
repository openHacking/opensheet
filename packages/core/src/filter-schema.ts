import { z } from 'zod';

const column = z.number().int().nonnegative();
const primitive = z.union([z.string().max(1000), z.number().finite(), z.boolean()]);
export const filterCriterionSchema = z.union([
  z.object({ column, operator: z.enum(['eq', 'neq']), value: primitive }).strict(),
  z
    .object({ column, operator: z.enum(['gt', 'gte', 'lt', 'lte']), value: z.number().finite() })
    .strict(),
  z
    .object({
      column,
      operator: z.enum(['contains', 'notContains', 'startsWith', 'endsWith']),
      value: z.string().max(1000),
    })
    .strict(),
  z
    .object({ column, operator: z.enum(['isBlank', 'isNotBlank', 'isError', 'isNotError']) })
    .strict(),
]);
export const filterSchema = z
  .union([
    z.object({ column, query: z.string().max(1000) }).strict(),
    z
      .object({
        mode: z.enum(['all', 'any']),
        headerRows: z.number().int().min(0).max(100000),
        criteria: z.array(filterCriterionSchema).min(1).max(64),
      })
      .strict(),
  ])
  .nullable();
export type WorksheetFilter = z.infer<typeof filterSchema>;
export type FilterCriterion = z.infer<typeof filterCriterionSchema>;
