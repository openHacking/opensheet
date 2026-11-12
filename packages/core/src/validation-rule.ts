import { z } from 'zod';
import type { CellValue } from './types.js';

// Internal building block only: no persisted schema or write enforcement is exposed yet.
const comparison = z.enum(['eq', 'neq', 'lt', 'lte', 'gt', 'gte', 'between', 'notBetween']);
const operand = z.number().finite();
const common = {
  allowBlank: z.boolean(),
  policy: z.enum(['reject', 'warn']),
};
const comparable = (kind: 'number' | 'date' | 'textLength') =>
  z
    .object({
      ...common,
      kind: z.literal(kind),
      operator: comparison,
      first: operand,
      second: operand.optional(),
    })
    .strict()
    .superRefine((rule, context) => {
      const needsSecond = rule.operator === 'between' || rule.operator === 'notBetween';
      if (needsSecond !== (rule.second !== undefined))
        context.addIssue({ code: 'custom', message: 'Range operators require exactly two bounds' });
      if (needsSecond && rule.second !== undefined && rule.first > rule.second)
        context.addIssue({ code: 'custom', message: 'Lower bound must not exceed upper bound' });
      if (
        kind === 'textLength' &&
        (!Number.isInteger(rule.first) ||
          rule.first < 0 ||
          (rule.second !== undefined && (!Number.isInteger(rule.second) || rule.second < 0)))
      )
        context.addIssue({
          code: 'custom',
          message: 'Text length bounds must be nonnegative integers',
        });
    });
export const validationRuleSchema = z.union([
  comparable('number'),
  comparable('date'),
  comparable('textLength'),
  z
    .object({ ...common, kind: z.literal('formula'), expression: z.string().min(1).max(4096) })
    .strict(),
  z
    .object({
      ...common,
      kind: z.literal('list'),
      values: z
        .array(z.union([z.string().max(32768), z.number().finite(), z.boolean()]))
        .min(1)
        .max(1000),
    })
    .strict()
    .refine(
      (rule) =>
        rule.values.reduce<number>(
          (total, value) => total + (typeof value === 'string' ? value.length : 0),
          0,
        ) <= 32768,
      'List text exceeds the rule budget',
    ),
]);
export type ValidationRule = z.infer<typeof validationRuleSchema>;

/** Typed comparisons; callers must supply a validated rule and calculated cell value. */
export function matchesValidationRule(rule: ValidationRule, value: CellValue): boolean {
  if (rule.kind === 'formula') return false;
  if (value === null || value === '') return rule.allowBlank;
  if (typeof value === 'object') return false;
  if (rule.kind === 'list') return rule.values.some((candidate) => candidate === value);
  const actual =
    rule.kind === 'textLength'
      ? typeof value === 'string'
        ? Array.from(value).length
        : undefined
      : typeof value === 'number' && Number.isFinite(value)
        ? value
        : undefined;
  if (actual === undefined) return false;
  switch (rule.operator) {
    case 'eq':
      return actual === rule.first;
    case 'neq':
      return actual !== rule.first;
    case 'lt':
      return actual < rule.first;
    case 'lte':
      return actual <= rule.first;
    case 'gt':
      return actual > rule.first;
    case 'gte':
      return actual >= rule.first;
    case 'between':
      return actual >= rule.first && actual <= rule.second!;
    case 'notBetween':
      return actual < rule.first || actual > rule.second!;
  }
}
