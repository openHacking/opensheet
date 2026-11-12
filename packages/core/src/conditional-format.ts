import { z } from 'zod';
import { parseFormula } from '@opensheetjs/formula';
import {
  assert,
  LIMITS,
  rectSchema,
  styleSchema,
  type CellStyle,
  type CellValue,
  type Rect,
} from './types.js';
import { transformValidationRange } from './validation-range.js';
import {
  rewriteValidationSheetReferences,
  transformValidationEntries,
} from './validation-model.js';
import { abortable, verificationYield } from './verification-yield.js';

// Internal model only. A persisted schema, commands, calculated-value plumbing and UI
// are prerequisites before conditional formatting can be exposed as a feature.
const numeric = z
  .object({
    kind: z.literal('number'),
    operator: z.enum(['eq', 'neq', 'lt', 'lte', 'gt', 'gte', 'between', 'notBetween']),
    first: z.number().finite(),
    second: z.number().finite().optional(),
  })
  .strict()
  .superRefine((value, context) => {
    const pair = value.operator === 'between' || value.operator === 'notBetween';
    if (
      pair !== (value.second !== undefined) ||
      (pair && value.second !== undefined && value.first > value.second)
    )
      context.addIssue({ code: 'custom', message: 'Invalid comparison bounds' });
  });
const textual = z
  .object({
    kind: z.literal('text'),
    operator: z.enum(['equals', 'notEquals', 'contains', 'notContains', 'startsWith', 'endsWith']),
    value: z.string().max(4096),
    caseSensitive: z.boolean(),
  })
  .strict();
const formula = z
  .object({ kind: z.literal('formula'), expression: z.string().min(1).max(4096) })
  .strict();
const visualStyle = styleSchema
  .pick({ bold: true, italic: true, underline: true, color: true, background: true })
  .strict();
const ruleSchema = z
  .object({
    id: z
      .string()
      .min(1)
      .max(100)
      .regex(/^[\w-]+$/)
      .refine((id) => !['__proto__', 'constructor', 'prototype'].includes(id)),
    range: rectSchema.strict(),
    condition: z.union([numeric, textual, formula]),
    style: visualStyle,
    stopIfTrue: z.boolean(),
  })
  .strict();
export type ConditionalRule = z.infer<typeof ruleSchema>;
const MAX_RULES = 1024;
const MAX_TEXT = 1_000_000;

/** Validate and detach rules; array order is priority, highest first. */
export function validateConditionalRules(
  input: unknown,
  rows: number,
  columns: number,
): ConditionalRule[] {
  assert(
    Number.isInteger(rows) &&
      rows > 0 &&
      rows <= LIMITS.rows &&
      Number.isInteger(columns) &&
      columns > 0 &&
      columns <= LIMITS.columns,
    'INVALID_ARGUMENT',
    'Invalid conditional formatting dimensions',
  );
  assert(
    Array.isArray(input) && input.length <= MAX_RULES,
    'LIMIT_EXCEEDED',
    'Too many conditional formatting rules',
  );
  const seen = new Set<string>();
  const validated: ConditionalRule[] = [];
  const expressions = new Set<string>();
  let text = 0;
  for (const item of input) {
    const parsed = ruleSchema.safeParse(item);
    assert(parsed.success, 'INVALID_ARGUMENT', 'Invalid conditional formatting rule');
    const rule = parsed.data;
    assert(!seen.has(rule.id), 'INVALID_ARGUMENT', 'Duplicate conditional formatting rule');
    seen.add(rule.id);
    assert(
      rule.range.startRow < rule.range.endRow &&
        rule.range.startColumn < rule.range.endColumn &&
        rule.range.endRow <= rows &&
        rule.range.endColumn <= columns,
      'INVALID_ARGUMENT',
      'Invalid conditional formatting range',
    );
    assert(
      Object.keys(rule.style).length > 0 &&
        Object.values(rule.style).every((value) => value !== undefined),
      'INVALID_ARGUMENT',
      'Conditional formatting style must not be empty',
    );
    if (rule.condition.kind === 'text') text += rule.condition.value.length;
    if (rule.condition.kind === 'formula') {
      text += rule.condition.expression.length;
      if (!expressions.has(rule.condition.expression)) {
        try {
          parseFormula(rule.condition.expression);
        } catch {
          assert(false, 'INVALID_ARGUMENT', 'Invalid conditional formatting formula');
        }
        expressions.add(rule.condition.expression);
      }
    }
    assert(text <= MAX_TEXT, 'LIMIT_EXCEEDED', 'Conditional formatting text budget exceeded');
    validated.push(rule);
  }
  return structuredClone(validated);
}

function matches(rule: ConditionalRule, value: CellValue): boolean {
  const condition = rule.condition;
  assert(
    condition.kind !== 'formula',
    'UNSUPPORTED_FEATURE',
    'Conditional formatting formula reader required',
  );
  if (condition.kind === 'number') {
    if (typeof value !== 'number' || !Number.isFinite(value)) return false;
    switch (condition.operator) {
      case 'eq':
        return value === condition.first;
      case 'neq':
        return value !== condition.first;
      case 'lt':
        return value < condition.first;
      case 'lte':
        return value <= condition.first;
      case 'gt':
        return value > condition.first;
      case 'gte':
        return value >= condition.first;
      case 'between':
        return value >= condition.first && value <= condition.second!;
      case 'notBetween':
        return value < condition.first || value > condition.second!;
    }
  }
  if (typeof value !== 'string') return false;
  const left = condition.caseSensitive ? value : value.toLowerCase();
  const right = condition.caseSensitive ? condition.value : condition.value.toLowerCase();
  switch (condition.operator) {
    case 'equals':
      return left === right;
    case 'notEquals':
      return left !== right;
    case 'contains':
      return left.includes(right);
    case 'notContains':
      return !left.includes(right);
    case 'startsWith':
      return left.startsWith(right);
    case 'endsWith':
      return left.endsWith(right);
  }
}

function applyVisual(overlay: CellStyle, style: ConditionalRule['style']) {
  if (overlay.bold === undefined && style.bold !== undefined) overlay.bold = style.bold;
  if (overlay.italic === undefined && style.italic !== undefined) overlay.italic = style.italic;
  if (overlay.underline === undefined && style.underline !== undefined)
    overlay.underline = style.underline;
  if (overlay.color === undefined && style.color !== undefined) overlay.color = style.color;
  if (overlay.background === undefined && style.background !== undefined)
    overlay.background = style.background;
}

/** Resolve only visual overrides without mutating the base style or rules. */
export function conditionalStyle(
  rules: readonly ConditionalRule[],
  row: number,
  column: number,
  value: CellValue,
  base: CellStyle,
): CellStyle {
  assert(
    !(value && typeof value === 'object' && value.error === '#LIMIT!'),
    'LIMIT_EXCEEDED',
    'Conditional formatting calculation exceeded its work budget',
  );
  assert(
    typeof value !== 'string' || value.length <= LIMITS.text,
    'LIMIT_EXCEEDED',
    'Conditional formatting value exceeds text budget',
  );
  assert(
    Number.isInteger(row) &&
      row >= 0 &&
      row < LIMITS.rows &&
      Number.isInteger(column) &&
      column >= 0 &&
      column < LIMITS.columns,
    'INVALID_ARGUMENT',
    'Invalid conditional formatting cell',
  );
  const overlay: CellStyle = {};
  for (const rule of rules) {
    const range = rule.range;
    if (
      row < range.startRow ||
      row >= range.endRow ||
      column < range.startColumn ||
      column >= range.endColumn ||
      !matches(rule, value)
    )
      continue;
    applyVisual(overlay, rule.style);
    if (rule.stopIfTrue) break;
  }
  return { ...base, ...overlay };
}

/** Evaluate a bounded set of already-loaded cells; callers supply calculated values. */
export function conditionalViewport(
  rules: readonly ConditionalRule[],
  cells: readonly { row: number; column: number; value: CellValue; base: CellStyle }[],
): CellStyle[] {
  checkViewportWork(rules, cells);
  assert(
    !rules.some((rule) => rule.condition.kind === 'formula'),
    'UNSUPPORTED_FEATURE',
    'Conditional formatting formula reader required',
  );
  return cells.map((cell) => conditionalStyle(rules, cell.row, cell.column, cell.value, cell.base));
}

type ViewCell = { row: number; column: number; value: CellValue; base: CellStyle };
type FormulaRequest = {
  row: number;
  column: number;
  expression: string;
  rowOffset: number;
  columnOffset: number;
};

function checkViewportWork(rules: readonly ConditionalRule[], cells: readonly ViewCell[]) {
  assert(
    Array.isArray(rules) &&
      rules.length <= MAX_RULES &&
      Array.isArray(cells) &&
      cells.length <= 10_000 &&
      rules.length * cells.length <= 1_000_000,
    'LIMIT_EXCEEDED',
    'Conditional formatting viewport exceeds work budget',
  );
  let textRules = 0;
  let ruleText = 0;
  for (const rule of rules) {
    if (rule.condition.kind === 'text') {
      textRules++;
      ruleText += rule.condition.value.length;
    } else if (rule.condition.kind === 'formula') ruleText += rule.condition.expression.length;
  }
  let valueText = 0;
  for (const cell of cells) {
    assert(
      !(cell.value && typeof cell.value === 'object' && cell.value.error === '#LIMIT!'),
      'LIMIT_EXCEEDED',
      'Conditional formatting calculation exceeded its work budget',
    );
    if (typeof cell.value !== 'string') continue;
    assert(
      cell.value.length <= LIMITS.text,
      'LIMIT_EXCEEDED',
      'Conditional formatting value exceeds text budget',
    );
    valueText += cell.value.length;
    assert(
      valueText <= 10_000_000,
      'LIMIT_EXCEEDED',
      'Conditional formatting viewport exceeds text work budget',
    );
  }
  assert(
    ruleText * cells.length + valueText * textRules <= 10_000_000,
    'LIMIT_EXCEEDED',
    'Conditional formatting viewport exceeds text work budget',
  );
}

/** Calculate against a caller-owned immutable revision. No partial result is returned on failure. */
export async function conditionalViewportWithFormulas(
  input: readonly ConditionalRule[],
  inputCells: readonly ViewCell[],
  evaluate: (request: FormulaRequest) => Promise<CellValue>,
  signal?: AbortSignal,
): Promise<CellStyle[]> {
  signal?.throwIfAborted();
  assert(
    Array.isArray(inputCells) && inputCells.length <= 10_000,
    'LIMIT_EXCEEDED',
    'Conditional formatting viewport exceeds work budget',
  );
  const rules = validateConditionalRules(input, LIMITS.rows, LIMITS.columns);
  checkViewportWork(rules, inputCells);
  signal?.throwIfAborted();
  const cells = structuredClone(inputCells);
  checkViewportWork(rules, cells);
  const checkpoint = verificationYield(signal);
  const results: CellStyle[] = [];
  let checks = 0;
  for (const cell of cells) {
    assert(
      Number.isInteger(cell.row) &&
        cell.row >= 0 &&
        cell.row < LIMITS.rows &&
        Number.isInteger(cell.column) &&
        cell.column >= 0 &&
        cell.column < LIMITS.columns,
      'INVALID_ARGUMENT',
      'Invalid conditional formatting cell',
    );
    const overlay: CellStyle = {};
    for (const rule of rules) {
      if (++checks % 1024 === 0) await checkpoint();
      signal?.throwIfAborted();
      const range = rule.range;
      if (
        cell.row < range.startRow ||
        cell.row >= range.endRow ||
        cell.column < range.startColumn ||
        cell.column >= range.endColumn
      )
        continue;
      let matched: boolean;
      if (rule.condition.kind === 'formula') {
        assert(
          typeof evaluate === 'function',
          'UNSUPPORTED_FEATURE',
          'Conditional formatting formula reader required',
        );
        const result = await abortable(
          evaluate({
            row: cell.row,
            column: cell.column,
            expression: rule.condition.expression,
            rowOffset: cell.row - range.startRow,
            columnOffset: cell.column - range.startColumn,
          }),
          signal,
        );
        signal?.throwIfAborted();
        assert(
          !(result && typeof result === 'object' && result.error === '#LIMIT!'),
          'LIMIT_EXCEEDED',
          'Conditional formatting calculation exceeded its work budget',
        );
        matched = result === true;
      } else matched = matches(rule, cell.value);
      if (!matched) continue;
      applyVisual(overlay, rule.style);
      if (rule.stopIfTrue) break;
    }
    results.push({ ...cell.base, ...overlay });
  }
  return results;
}

/** Adjust ranges for a structural edit; deleted ranges are removed. */
export function transformConditionalRules(
  input: unknown,
  rows: number,
  columns: number,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  index: number,
  count: number,
  sheetName?: string,
  editedSheet?: { name: string; axisLength: number },
): ConditionalRule[] {
  const rules = validateConditionalRules(input, rows, columns);
  if (rules.some((rule) => rule.condition.kind === 'formula') || editedSheet) {
    const entries = rules.map((rule) => ({
      id: rule.id,
      range: rule.range,
      rule:
        rule.condition.kind === 'formula'
          ? {
              kind: 'formula',
              expression: rule.condition.expression,
              allowBlank: false,
              policy: 'reject',
            }
          : { kind: 'number', operator: 'eq', first: 0, allowBlank: false, policy: 'reject' },
    }));
    const transformed = transformValidationEntries(
      entries,
      rows,
      columns,
      axis,
      operation,
      index,
      count,
      sheetName,
      editedSheet,
    );
    const byId = new Map(rules.map((rule) => [rule.id, rule]));
    const output = transformed.map((entry) => {
      const original = byId.get(entry.id)!;
      return {
        ...original,
        range: entry.range,
        condition:
          original.condition.kind === 'formula' && entry.rule.kind === 'formula'
            ? { kind: 'formula' as const, expression: entry.rule.expression }
            : original.condition,
      };
    });
    const local = !editedSheet || sheetName?.toLowerCase() === editedSheet.name.toLowerCase();
    const delta = local ? (operation === 'insert' ? count : -count) : 0;
    return validateConditionalRules(
      output,
      rows + (axis === 'row' ? delta : 0),
      columns + (axis === 'column' ? delta : 0),
    );
  }
  const axisLength = axis === 'row' ? rows : columns;
  // Check the operation even for an empty collection.
  transformValidationRange(
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    axis,
    operation,
    index,
    count,
    axisLength,
  );
  const transformed = rules.flatMap((rule) => {
    const range: Rect | null = transformValidationRange(
      rule.range,
      axis,
      operation,
      index,
      count,
      axisLength,
    );
    return range ? [{ ...rule, range }] : [];
  });
  const delta = operation === 'insert' ? count : -count;
  return validateConditionalRules(
    transformed,
    rows + (axis === 'row' ? delta : 0),
    columns + (axis === 'column' ? delta : 0),
  );
}

/** Rewrite explicit worksheet references without changing rule order or visual styles. */
export function rewriteConditionalSheetReferences(
  input: unknown,
  rows: number,
  columns: number,
  oldName: string,
  newName: string | null,
): ConditionalRule[] {
  const rules = validateConditionalRules(input, rows, columns);
  const entries = rules.flatMap((rule) =>
    rule.condition.kind === 'formula'
      ? [
          {
            id: rule.id,
            range: rule.range,
            rule: {
              kind: 'formula',
              expression: rule.condition.expression,
              allowBlank: false,
              policy: 'reject',
            },
          },
        ]
      : [],
  );
  const rewritten = rewriteValidationSheetReferences(entries, rows, columns, oldName, newName);
  const expressions = new Map(rewritten.map((entry) => [entry.id, entry.rule]));
  for (const rule of rules) {
    if (rule.condition.kind !== 'formula') continue;
    const rewrittenRule = expressions.get(rule.id);
    assert(rewrittenRule?.kind === 'formula', 'INVALID_ARGUMENT', 'Missing rewritten formula');
    rule.condition.expression = rewrittenRule.expression;
  }
  return validateConditionalRules(rules, rows, columns);
}
