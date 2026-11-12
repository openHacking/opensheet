import { validateRect } from './address.js';
import { validateValidationEntries, type ValidationEntry } from './validation-model.js';
import { assert, type Rect } from './types.js';
import { verificationYield } from './verification-yield.js';

const MAX_RANGES = 65536;
const MAX_CANDIDATES = 65536;
const MAX_WORK = 1_000_000;

/** Collect direct intersections only; callers must separately include formula dependents. */
export async function validationCandidates(
  inputRules: readonly ValidationEntry[],
  inputRanges: readonly Rect[],
  dimensions: { rows: number; columns: number },
  signal?: AbortSignal,
) {
  signal?.throwIfAborted();
  assert(inputRanges.length <= MAX_RANGES, 'LIMIT_EXCEEDED', 'Too many validation ranges');
  const { rows, columns } = dimensions;
  const rules = validateValidationEntries(inputRules, rows, columns);
  const ranges = inputRanges.map((range) => {
    validateRect(range);
    assert(
      range.endRow <= rows && range.endColumn <= columns,
      'INVALID_RANGE',
      'Range exceeds sheet dimensions',
    );
    return { ...range };
  });
  assert(
    rules.length * ranges.length <= MAX_WORK,
    'LIMIT_EXCEEDED',
    'Validation candidate work exceeded',
  );
  const checkpoint = verificationYield(signal);
  const positions = new Map<number, { row: number; column: number }>();
  let work = rules.length * ranges.length;
  let scanned = 0;
  for (const range of ranges) {
    for (const { range: rule } of rules) {
      if (++scanned % 1024 === 0) await checkpoint();
      const startRow = Math.max(range.startRow, rule.startRow);
      const endRow = Math.min(range.endRow, rule.endRow);
      const startColumn = Math.max(range.startColumn, rule.startColumn);
      const endColumn = Math.min(range.endColumn, rule.endColumn);
      for (let row = startRow; row < endRow; row++) {
        for (let column = startColumn; column < endColumn; column++) {
          assert(++work <= MAX_WORK, 'LIMIT_EXCEEDED', 'Validation candidate work exceeded');
          if (++scanned % 1024 === 0) await checkpoint();
          const key = row * columns + column;
          if (!positions.has(key)) {
            assert(
              positions.size < MAX_CANDIDATES,
              'LIMIT_EXCEEDED',
              'Too many validation candidates',
            );
            positions.set(key, { row, column });
          }
        }
      }
    }
  }
  signal?.throwIfAborted();
  return {
    rules,
    positions: [...positions.values()].sort((a, b) => a.row - b.row || a.column - b.column),
  };
}
