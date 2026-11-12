import { assert, LIMITS, type CellValue } from './types.js';
import { matchesValidationRule } from './validation-rule.js';
import { validateValidationEntries, type ValidationEntry } from './validation-model.js';
import { abortable, verificationYield } from './verification-yield.js';

// Provisional internal bounds; not part of the public workbook contract.
const MAX_RULES = 1024;
const MAX_CELLS = 65536;
const MAX_WORK = 1_000_000;
const MAX_DETAILS = 32;
const CHECKPOINT = Symbol('validation checkpoint');
const CHECKPOINT_INTERVAL = 1024;

type Rule = ValidationEntry;
type Cell = { row: number; column: number; value: CellValue };
type Position = Pick<Cell, 'row' | 'column'>;
type Violation = { row: number; column: number; ruleId: string; policy: 'reject' | 'warn' };
type Report = { rejected: number; warnings: number; details: Violation[]; truncated: boolean };
type FormulaRequest = Position & { expression: string; rowOffset: number; columnOffset: number };
type Dimensions = { rows: number; columns: number };

/** Pure, bounded prototype. This does not publish, approve warnings or enforce commands. */
function* checkPositions(
  rules: readonly Rule[],
  cells: readonly Position[],
  dimensions: Dimensions = { rows: LIMITS.rows, columns: LIMITS.columns },
): Generator<number | FormulaRequest | typeof CHECKPOINT, Report, CellValue> {
  assert(
    rules.length <= MAX_RULES && cells.length <= MAX_CELLS,
    'LIMIT_EXCEEDED',
    'Validation batch exceeds limits',
  );
  assert(
    rules.length * cells.length <= MAX_WORK,
    'LIMIT_EXCEEDED',
    'Validation work budget exceeded',
  );
  const checkedRules = validateValidationEntries(rules, dimensions.rows, dimensions.columns);
  const lists: Array<ReadonlySet<string | number | boolean> | undefined> = [];
  const bounds = { startRow: Infinity, endRow: 0, startColumn: Infinity, endColumn: 0 };
  let work = rules.length * cells.length;
  for (const entry of checkedRules) {
    bounds.startRow = Math.min(bounds.startRow, entry.range.startRow);
    bounds.endRow = Math.max(bounds.endRow, entry.range.endRow);
    bounds.startColumn = Math.min(bounds.startColumn, entry.range.startColumn);
    bounds.endColumn = Math.max(bounds.endColumn, entry.range.endColumn);
    const weight =
      entry.rule.kind === 'list'
        ? entry.rule.values.reduce<number>(
            (sum, value) => sum + (typeof value === 'string' ? value.length + 1 : 1),
            0,
          )
        : entry.rule.kind === 'formula'
          ? entry.rule.expression.length
          : 1;
    work += weight;
    assert(work <= MAX_WORK, 'LIMIT_EXCEEDED', 'Validation work budget exceeded');
    lists.push(entry.rule.kind === 'list' ? new Set(entry.rule.values) : undefined);
  }
  for (const [index, cell] of cells.entries()) {
    if (index > 0 && index % CHECKPOINT_INTERVAL === 0) yield CHECKPOINT;
    assert(
      Number.isInteger(cell.row) &&
        cell.row >= 0 &&
        cell.row < dimensions.rows &&
        Number.isInteger(cell.column) &&
        cell.column >= 0 &&
        cell.column < dimensions.columns,
      'INVALID_ARGUMENT',
      'Invalid validation cell position',
    );
  }
  let rejected = 0,
    warnings = 0;
  const details: Violation[] = [];
  let scanned = 0;
  for (const [cellIndex, cell] of cells.entries()) {
    if (
      cell.row < bounds.startRow ||
      cell.row >= bounds.endRow ||
      cell.column < bounds.startColumn ||
      cell.column >= bounds.endColumn
    ) {
      const previous = Math.floor(scanned / CHECKPOINT_INTERVAL);
      scanned += checkedRules.length;
      if (Math.floor(scanned / CHECKPOINT_INTERVAL) !== previous) yield CHECKPOINT;
      continue;
    }
    let value: CellValue | undefined;
    for (const [index, { id, range, rule }] of checkedRules.entries()) {
      if (++scanned % CHECKPOINT_INTERVAL === 0) yield CHECKPOINT;
      if (
        cell.row < range.startRow ||
        cell.row >= range.endRow ||
        cell.column < range.startColumn ||
        cell.column >= range.endColumn
      )
        continue;
      if (value === undefined) value = yield cellIndex;
      assert(
        !(value && typeof value === 'object' && value.error === '#LIMIT!'),
        'LIMIT_EXCEEDED',
        'Validation calculation exceeded its work budget',
      );
      if ((rule.kind === 'list' || rule.kind === 'textLength') && typeof value === 'string')
        work += value.length;
      assert(work <= MAX_WORK, 'LIMIT_EXCEEDED', 'Validation work budget exceeded');
      const list = lists[index];
      let matches: boolean;
      if (value !== null && typeof value === 'object') {
        // A rule cannot approve an error in the cell being validated.
        matches = false;
      } else if (rule.kind === 'formula') {
        if (value === null || value === '') matches = rule.allowBlank;
        else {
          work += rule.expression.length;
          assert(work <= MAX_WORK, 'LIMIT_EXCEEDED', 'Validation work budget exceeded');
          const result = yield {
            row: cell.row,
            column: cell.column,
            expression: rule.expression,
            rowOffset: cell.row - range.startRow,
            columnOffset: cell.column - range.startColumn,
          };
          assert(
            !(result && typeof result === 'object' && result.error === '#LIMIT!'),
            'LIMIT_EXCEEDED',
            'Validation calculation exceeded its work budget',
          );
          matches = result === true;
        }
      } else
        matches = list
          ? value === null || value === ''
            ? rule.allowBlank
            : typeof value !== 'object' && list.has(value)
          : matchesValidationRule(rule, value);
      if (matches) continue;
      if (rule.policy === 'reject') rejected++;
      else warnings++;
      if (details.length < MAX_DETAILS)
        details.push({ row: cell.row, column: cell.column, ruleId: id, policy: rule.policy });
    }
  }
  return { rejected, warnings, details, truncated: rejected + warnings > details.length };
}

export function checkValidationBatch(rules: readonly Rule[], cells: readonly Cell[]) {
  const checks = checkPositions(rules, cells);
  let step = checks.next();
  while (!step.done) {
    if (step.value === CHECKPOINT) {
      step = checks.next();
      continue;
    }
    assert(
      typeof step.value === 'number',
      'UNSUPPORTED_FEATURE',
      'Validation formula reader required',
    );
    step = checks.next(cells[step.value].value);
  }
  return step.value;
}

/** Read covered cells from a caller-owned, stable staged revision. Never publishes a write. */
export async function checkStagedValidation(
  rules: readonly Rule[],
  positions: readonly Position[],
  read: (row: number, column: number) => Promise<CellValue>,
  signal?: AbortSignal,
  evaluateFormula?: (request: FormulaRequest) => Promise<CellValue>,
  dimensions?: Dimensions,
): Promise<Report> {
  const checkpoint = verificationYield(signal);
  signal?.throwIfAborted();
  assert(positions.length <= MAX_CELLS, 'LIMIT_EXCEEDED', 'Validation batch exceeds limits');
  const snapshot = positions.map(({ row, column }) => ({ row, column }));
  const limits = dimensions ? { ...dimensions } : undefined;
  const checks = checkPositions(rules, snapshot, limits);
  let step = checks.next();
  while (!step.done) {
    await checkpoint();
    signal?.throwIfAborted();
    if (step.value === CHECKPOINT) {
      step = checks.next();
      continue;
    }
    let value: CellValue;
    if (typeof step.value === 'number') {
      const { row, column } = snapshot[step.value];
      value = await abortable(read(row, column), signal);
    } else {
      assert(evaluateFormula, 'UNSUPPORTED_FEATURE', 'Validation formula reader required');
      value = await abortable(evaluateFormula(step.value), signal);
    }
    signal?.throwIfAborted();
    step = checks.next(value);
  }
  return step.value;
}
