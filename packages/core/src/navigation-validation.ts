import { assert } from './types.js';

/** Check the same untrusted request before dispatch and at the Worker boundary. */
export function validateNavigationArguments(
  sheet: unknown,
  row: unknown,
  column: unknown,
  axis: unknown,
  direction: unknown,
) {
  assert(
    typeof sheet === 'string' && sheet.length > 0,
    'INVALID_ARGUMENT',
    'Invalid navigation sheet',
  );
  assert(
    Number.isSafeInteger(row) &&
      (row as number) >= 0 &&
      Number.isSafeInteger(column) &&
      (column as number) >= 0,
    'INVALID_ARGUMENT',
    'Invalid navigation position',
  );
  assert(
    (axis === 'row' || axis === 'column') && (direction === -1 || direction === 1),
    'INVALID_ARGUMENT',
    'Invalid navigation direction',
  );
}
