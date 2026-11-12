import { assert } from './types.js';

/** Validate before sending to a Worker and again at the Worker boundary. */
export function validateSearchArguments(
  sheet: unknown,
  query: unknown,
  after: unknown,
  direction: unknown,
) {
  assert(typeof sheet === 'string' && sheet.length > 0, 'INVALID_ARGUMENT', 'Invalid search sheet');
  assert(typeof query === 'string', 'INVALID_ARGUMENT', 'Search query must be text');
  assert(query.length <= 65_536, 'LIMIT_EXCEEDED', 'Search query exceeds 65,536 UTF-16 units');
  assert(
    Number.isSafeInteger(after) && (after as number) >= -1,
    'INVALID_ARGUMENT',
    'Invalid search position',
  );
  assert(
    direction === undefined || direction === 'next' || direction === 'previous',
    'INVALID_ARGUMENT',
    'Invalid search direction',
  );
}
