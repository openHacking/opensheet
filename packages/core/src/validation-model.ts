import { assert, LIMITS, rectSchema, type Rect } from './types.js';
import { validationRuleSchema, type ValidationRule } from './validation-rule.js';
import { transformValidationRange } from './validation-range.js';
import { transformDeletedFormulaRange } from './deleted-formula-range.js';
import {
  offsetFormula,
  parseFormula,
  printFormula,
  transformReferences,
} from '@opensheetjs/formula';

/** Internal staged model. Mandatory validation is not yet persisted or exposed. */
export type ValidationEntry = { id: string; range: Rect; rule: ValidationRule };
const MAX_RULES = 1024;
const MAX_METADATA_WORK = 1_000_000;

export function validateValidationEntries(
  input: unknown,
  rows: number,
  columns: number,
): ValidationEntry[] {
  assert(
    Number.isInteger(rows) &&
      rows > 0 &&
      rows <= LIMITS.rows &&
      Number.isInteger(columns) &&
      columns > 0 &&
      columns <= LIMITS.columns,
    'INVALID_ARGUMENT',
    'Invalid validation dimensions',
  );
  assert(
    Array.isArray(input) && input.length <= MAX_RULES,
    'LIMIT_EXCEEDED',
    'Too many validation rules',
  );
  const ids = new Set<string>();
  const parsedFormulas = new Set<string>();
  let work = input.length;
  for (const item of input) {
    assert(
      item &&
        typeof item === 'object' &&
        !Array.isArray(item) &&
        Object.keys(item).every((key) => ['id', 'range', 'rule'].includes(key)),
      'INVALID_ARGUMENT',
      'Invalid validation entry',
    );
    const { id, range, rule } = item as Record<string, any>;
    assert(
      typeof id === 'string' && id.length > 0 && id.length <= 100 && !ids.has(id),
      'INVALID_ARGUMENT',
      'Invalid validation rule range or identity',
    );
    ids.add(id);
    assert(
      rectSchema.strict().safeParse(range).success &&
        range.endRow > range.startRow &&
        range.endColumn > range.startColumn &&
        range.endRow <= rows &&
        range.endColumn <= columns,
      'INVALID_ARGUMENT',
      'Invalid validation rule range or identity',
    );
    if (rule?.kind === 'list' && Array.isArray(rule.values)) {
      assert(rule.values.length <= 1000, 'LIMIT_EXCEEDED', 'Validation list exceeds item budget');
      for (const value of rule.values) {
        work += typeof value === 'string' ? value.length + 1 : 1;
        assert(work <= MAX_METADATA_WORK, 'LIMIT_EXCEEDED', 'Validation work budget exceeded');
      }
    }
    if (rule?.kind === 'formula' && typeof rule.expression === 'string') {
      work += rule.expression.length;
      assert(work <= MAX_METADATA_WORK, 'LIMIT_EXCEEDED', 'Validation work budget exceeded');
    }
    assert(
      validationRuleSchema.safeParse(rule).success,
      'INVALID_ARGUMENT',
      'Invalid validation rule',
    );
    if (rule.kind === 'formula' && !parsedFormulas.has(rule.expression)) {
      try {
        parseFormula(rule.expression);
      } catch {
        assert(false, 'INVALID_ARGUMENT', 'Invalid validation formula syntax');
      }
      parsedFormulas.add(rule.expression);
    }
  }
  return structuredClone(input) as ValidationEntry[];
}

/** Update explicit references after a worksheet rename or removal. */
export function rewriteValidationSheetReferences(
  input: unknown,
  rows: number,
  columns: number,
  oldName: string,
  newName: string | null,
): ValidationEntry[] {
  assert(
    oldName && (newName === null || newName),
    'INVALID_ARGUMENT',
    'Missing validation sheet name',
  );
  const name = oldName.toLowerCase();
  const entries = validateValidationEntries(input, rows, columns);
  for (const { rule } of entries) {
    if (rule.kind !== 'formula') continue;
    try {
      let changed = false;
      const ast = transformReferences(parseFormula(rule.expression), (ref) => {
        if (ref.sheet?.toLowerCase() !== name) return ref;
        changed = true;
        return newName === null ? null : { ...ref, sheet: newName };
      });
      if (changed) rule.expression = printFormula(ast);
    } catch {
      assert(false, 'UNSUPPORTED_FEATURE', 'Validation formula cannot safely be rewritten');
    }
  }
  return validateValidationEntries(entries, rows, columns);
}

/** Apply insertion/deletion to every range while retaining rule identity and order. */
export function transformValidationEntries(
  input: unknown,
  rows: number,
  columns: number,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  index: number,
  count: number,
  sheetName?: string,
  editedSheet?: { name: string; axisLength: number },
): ValidationEntry[] {
  const entries = validateValidationEntries(input, rows, columns);
  assert(
    !editedSheet || (sheetName && editedSheet.name),
    'INVALID_ARGUMENT',
    'Missing validation sheet name',
  );
  const targetName = (editedSheet?.name ?? sheetName)?.toLowerCase();
  const local = !editedSheet || sheetName?.toLowerCase() === targetName;
  const axisLength = local ? (axis === 'row' ? rows : columns) : editedSheet!.axisLength;
  // Validate the operation even when the collection is empty.
  const sample: Rect = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
  transformValidationRange(sample, axis, operation, index, count, axisLength);
  const transformed = entries.flatMap((entry) => {
    const range = local
      ? transformValidationRange(entry.range, axis, operation, index, count, axisLength)
      : entry.range;
    if (!range) return [];
    if (entry.rule.kind !== 'formula') return [{ ...entry, range }];
    assert(sheetName, 'INVALID_ARGUMENT', 'Validation formula rewrite requires a sheet name');
    const coordinate = axis === 'row' ? 'row' : 'column';
    const origin = axis === 'row' ? entry.range.startRow : entry.range.startColumn;
    // Re-anchor to the first surviving member if the old origin was removed.
    const delta =
      local && operation === 'delete' && origin >= index && origin < index + count
        ? index + count - origin
        : 0;
    let expression: string;
    try {
      const source = delta
        ? offsetFormula(
            entry.rule.expression,
            axis === 'row' ? delta : 0,
            axis === 'column' ? delta : 0,
          )
        : entry.rule.expression;
      let changed = delta !== 0;
      const rewriteReference: Parameters<typeof transformReferences>[1] = (ref) => {
        if ((ref.sheet ?? sheetName).toLowerCase() !== targetName) return ref;
        const position = ref[coordinate];
        if (position < index) return ref;
        changed = true;
        if (operation === 'insert') return { ...ref, [coordinate]: position + count };
        return position < index + count ? null : { ...ref, [coordinate]: position - count };
      };
      const rewritten = transformReferences(
        parseFormula(source),
        rewriteReference,
        operation === 'delete'
          ? (start, end) => {
              if (
                (start.sheet ?? sheetName).toLowerCase() !== targetName ||
                (end.sheet ?? sheetName).toLowerCase() !== targetName
              ) {
                const first = rewriteReference(start),
                  last = rewriteReference(end);
                return first && last ? [first, last] : null;
              }
              const range = transformDeletedFormulaRange(start, end, axis, index, count);
              if (
                !range ||
                range[0][coordinate] !== start[coordinate] ||
                range[1][coordinate] !== end[coordinate]
              )
                changed = true;
              return range;
            }
          : undefined,
      );
      expression = changed ? printFormula(rewritten) : source;
    } catch {
      assert(false, 'UNSUPPORTED_FEATURE', 'Validation formula cannot safely be rewritten');
    }
    return [{ ...entry, range, rule: { ...entry.rule, expression } }];
  });
  const delta = local ? (operation === 'insert' ? count : -count) : 0;
  return validateValidationEntries(
    transformed,
    rows + (axis === 'row' ? delta : 0),
    columns + (axis === 'column' ? delta : 0),
  );
}
