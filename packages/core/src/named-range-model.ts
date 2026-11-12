import { assert, LIMITS, rectSchema, type Rect } from './types.js';
import { transformValidationRange } from './validation-range.js';
import { verificationYield } from './verification-yield.js';

// Internal rectangle-only model. It is not a persisted schema or a public feature.
export type NamedRangeSheet = { id: string; rows: number; columns: number };
export type NamedRange = {
  id: string;
  name: string;
  scopeSheetId: string | null;
  target: { sheetId: string; range: Rect } | null;
};

const MAX_NAMES = 8192;
const MAX_NAME_TEXT = 1_000_000;
const identifier = /^[A-Za-z_][A-Za-z_0-9.]*$/;
const safeId = /^(?!__proto__$|constructor$|prototype$)[\w-]{1,100}$/;
const strictRectSchema = rectSchema.strict();

function validName(name: unknown): name is string {
  return (
    typeof name === 'string' &&
    name.length <= 255 &&
    identifier.test(name) &&
    !/^[A-Za-z]+[1-9]\d*/.test(name) &&
    !/^(?:TRUE|FALSE|R|C)$/i.test(name)
  );
}

function sheetMap(sheets: readonly NamedRangeSheet[]) {
  assert(
    Array.isArray(sheets) && sheets.length > 0 && sheets.length <= LIMITS.sheets,
    'INVALID_ARGUMENT',
    'Invalid named-range sheets',
  );
  const result = new Map<string, NamedRangeSheet>();
  for (const sheet of sheets) {
    assert(
      sheet &&
        typeof sheet === 'object' &&
        !Array.isArray(sheet) &&
        ['id', 'rows', 'columns'].every((key) => Object.hasOwn(sheet, key)) &&
        typeof sheet.id === 'string' &&
        safeId.test(sheet.id) &&
        !result.has(sheet.id) &&
        Number.isInteger(sheet.rows) &&
        sheet.rows > 0 &&
        sheet.rows <= LIMITS.rows &&
        Number.isInteger(sheet.columns) &&
        sheet.columns > 0 &&
        sheet.columns <= LIMITS.columns,
      'INVALID_ARGUMENT',
      'Invalid named-range sheet',
    );
    result.set(sheet.id, sheet);
  }
  return result;
}

/** Validate identities, scope and rectangle geometry and return detached data. */
function* validateNamedRangesSteps(
  input: unknown,
  sheets: readonly NamedRangeSheet[],
): Generator<void, NamedRange[]> {
  const bySheet = sheetMap(sheets);
  assert(Array.isArray(input), 'INVALID_ARGUMENT', 'Invalid named-range collection');
  assert(input.length <= MAX_NAMES, 'LIMIT_EXCEEDED', 'Too many named ranges');
  const ids = new Set<string>(),
    scopes = new Map<string | null, Set<string>>();
  let text = 0;
  const output: NamedRange[] = [];
  for (const [position, entry] of input.entries()) {
    if (position > 0 && position % 64 === 0) yield;
    assert(
      entry &&
        typeof entry === 'object' &&
        !Array.isArray(entry) &&
        ['id', 'name', 'scopeSheetId', 'target'].every((key) => Object.hasOwn(entry, key)) &&
        Object.keys(entry).every((key) => ['id', 'name', 'scopeSheetId', 'target'].includes(key)),
      'INVALID_ARGUMENT',
      'Invalid named range',
    );
    const { id, name, scopeSheetId, target } = entry as Record<string, any>;
    assert(
      typeof id === 'string' && safeId.test(id) && !ids.has(id) && validName(name),
      'INVALID_ARGUMENT',
      'Invalid named-range identity or name',
    );
    assert(
      scopeSheetId === null || (typeof scopeSheetId === 'string' && bySheet.has(scopeSheetId)),
      'INVALID_ARGUMENT',
      'Invalid named-range scope',
    );
    const scope = scopes.get(scopeSheetId) ?? new Set<string>();
    const key = name.toUpperCase();
    assert(!scope.has(key), 'INVALID_ARGUMENT', 'Duplicate named range in scope');
    scope.add(key);
    scopes.set(scopeSheetId, scope);
    ids.add(id);
    text += name.length;
    assert(text <= MAX_NAME_TEXT, 'LIMIT_EXCEEDED', 'Named-range metadata budget exceeded');
    if (target !== null) {
      assert(
        target &&
          typeof target === 'object' &&
          !Array.isArray(target) &&
          ['sheetId', 'range'].every((key) => Object.hasOwn(target, key)) &&
          Object.keys(target).every((key) => ['sheetId', 'range'].includes(key)) &&
          typeof target.sheetId === 'string' &&
          bySheet.has(target.sheetId),
        'INVALID_ARGUMENT',
        'Invalid named-range target',
      );
      const sheet = bySheet.get(target.sheetId)!;
      assert(
        target.range &&
          typeof target.range === 'object' &&
          !Array.isArray(target.range) &&
          ['startRow', 'startColumn', 'endRow', 'endColumn'].every((key) =>
            Object.hasOwn(target.range, key),
          ),
        'INVALID_ARGUMENT',
        'Invalid named-range target',
      );
      const parsed = strictRectSchema.safeParse(target.range);
      assert(
        parsed.success &&
          parsed.data.startRow < parsed.data.endRow &&
          parsed.data.startColumn < parsed.data.endColumn &&
          parsed.data.endRow <= sheet.rows &&
          parsed.data.endColumn <= sheet.columns,
        'INVALID_ARGUMENT',
        'Invalid named-range target',
      );
      output.push({
        id,
        name,
        scopeSheetId,
        target: { sheetId: target.sheetId, range: parsed.data },
      });
    } else output.push({ id, name, scopeSheetId, target: null });
  }
  return output;
}

/** Validate identities, scope and rectangle geometry and return detached data. */
export function validateNamedRanges(
  input: unknown,
  sheets: readonly NamedRangeSheet[],
): NamedRange[] {
  const steps = validateNamedRangesSteps(input, sheets);
  let step = steps.next();
  while (!step.done) step = steps.next();
  return step.value;
}

/** Internal staged-revision API: callers must keep input and sheets stable until it settles. */
export async function validateStagedNamedRanges(
  input: unknown,
  sheets: readonly NamedRangeSheet[],
  signal?: AbortSignal,
): Promise<NamedRange[]> {
  signal?.throwIfAborted();
  const checkpoint = verificationYield(signal);
  const steps = validateNamedRangesSteps(input, sheets);
  let step = steps.next();
  while (!step.done) {
    await checkpoint();
    step = steps.next();
  }
  signal?.throwIfAborted();
  return step.value;
}

/** Build a lookup once per immutable metadata revision; local names shadow workbook names. */
export function indexNamedRanges(input: unknown, sheets: readonly NamedRangeSheet[]) {
  const validated = validateNamedRanges(input, sheets),
    bySheet = sheetMap(sheets);
  const globals = new Map<string, NamedRange>(),
    locals = new Map<string, Map<string, NamedRange>>();
  for (const entry of validated) {
    const key = entry.name.toUpperCase();
    if (entry.scopeSheetId === null) globals.set(key, entry);
    else {
      const scope = locals.get(entry.scopeSheetId) ?? new Map<string, NamedRange>();
      scope.set(key, entry);
      locals.set(entry.scopeSheetId, scope);
    }
  }
  return {
    resolve(name: string, contextSheetId: string): NamedRange | undefined {
      assert(
        typeof name === 'string' && validName(name) && bySheet.has(contextSheetId),
        'INVALID_ARGUMENT',
        'Invalid named-range lookup',
      );
      const key = name.toUpperCase();
      const found = locals.get(contextSheetId)?.get(key) ?? globals.get(key);
      return found ? structuredClone(found) : undefined;
    },
  };
}

/** Transform target geometry against one pre-edit sheet; deleted targets remain broken. */
export function transformNamedRanges(
  input: unknown,
  sheets: readonly NamedRangeSheet[],
  sheetId: string,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  index: number,
  count: number,
): NamedRange[] {
  const bySheet = sheetMap(sheets),
    sheet = bySheet.get(sheetId);
  assert(sheet, 'INVALID_ARGUMENT', 'Unknown named-range target sheet');
  const entries = validateNamedRanges(input, sheets);
  const length = axis === 'row' ? sheet.rows : sheet.columns;
  transformValidationRange(
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    axis,
    operation,
    index,
    count,
    length,
  );
  for (const entry of entries) {
    if (entry.target?.sheetId !== sheetId) continue;
    const range = transformValidationRange(
      entry.target.range,
      axis,
      operation,
      index,
      count,
      length,
    );
    entry.target = range ? { sheetId, range } : null;
  }
  const updated = sheets.map((item) =>
    item.id === sheetId
      ? {
          ...item,
          [axis === 'row' ? 'rows' : 'columns']: length + (operation === 'insert' ? count : -count),
        }
      : item,
  );
  return validateNamedRanges(entries, updated);
}

/** Transform an immutable staged revision without publishing partial results. */
export async function transformStagedNamedRanges(
  input: unknown,
  sheets: readonly NamedRangeSheet[],
  sheetId: string,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  index: number,
  count: number,
  signal?: AbortSignal,
): Promise<NamedRange[]> {
  signal?.throwIfAborted();
  const sheet = sheetMap(sheets).get(sheetId);
  assert(sheet, 'INVALID_ARGUMENT', 'Unknown named-range target sheet');
  const entries = await validateStagedNamedRanges(input, sheets, signal);
  const length = axis === 'row' ? sheet.rows : sheet.columns;
  transformValidationRange(
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    axis,
    operation,
    index,
    count,
    length,
  );
  const checkpoint = verificationYield(signal);
  for (const [position, entry] of entries.entries()) {
    if (position > 0 && position % 64 === 0) await checkpoint();
    if (entry.target?.sheetId !== sheetId) continue;
    const range = transformValidationRange(
      entry.target.range,
      axis,
      operation,
      index,
      count,
      length,
    );
    entry.target = range ? { sheetId, range } : null;
  }
  const updated = sheets.map((item) =>
    item.id === sheetId
      ? {
          ...item,
          [axis === 'row' ? 'rows' : 'columns']: length + (operation === 'insert' ? count : -count),
        }
      : item,
  );
  return validateStagedNamedRanges(entries, updated, signal);
}

/** Drop names scoped to a removed sheet and break surviving references to that sheet. */
export function removeNamedRangeSheet(
  input: unknown,
  sheets: readonly NamedRangeSheet[],
  sheetId: string,
): NamedRange[] {
  const bySheet = sheetMap(sheets);
  assert(
    bySheet.has(sheetId) && bySheet.size > 1,
    'INVALID_ARGUMENT',
    'Cannot remove named-range sheet',
  );
  const entries = validateNamedRanges(input, sheets).filter(
    (entry) => entry.scopeSheetId !== sheetId,
  );
  for (const entry of entries) if (entry.target?.sheetId === sheetId) entry.target = null;
  return validateNamedRanges(
    entries,
    sheets.filter((sheet) => sheet.id !== sheetId),
  );
}
