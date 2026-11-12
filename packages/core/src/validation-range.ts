import { assert, LIMITS, rectSchema, type Rect } from './types.js';

/** Internal prototype for half-open rule ranges; not integrated with persisted rules. */
export function transformValidationRange(
  range: Rect,
  axis: 'row' | 'column',
  operation: 'insert' | 'delete',
  index: number,
  count: number,
  axisLength: number,
): Rect | null {
  const max = axis === 'row' ? LIMITS.rows : LIMITS.columns;
  const startKey = axis === 'row' ? 'startRow' : 'startColumn';
  const endKey = axis === 'row' ? 'endRow' : 'endColumn';
  assert(
    rectSchema.safeParse(range).success &&
      range.endRow > range.startRow &&
      range.endColumn > range.startColumn &&
      range.endRow <= LIMITS.rows &&
      range.endColumn <= LIMITS.columns &&
      (axis === 'row' || axis === 'column') &&
      (operation === 'insert' || operation === 'delete') &&
      Number.isInteger(axisLength) &&
      axisLength > 0 &&
      axisLength <= max &&
      range[endKey] <= axisLength &&
      Number.isInteger(index) &&
      index >= 0 &&
      Number.isInteger(count) &&
      count > 0 &&
      (operation === 'insert'
        ? index <= axisLength && axisLength + count <= max
        : index + count <= axisLength && axisLength - count >= 1),
    'INVALID_ARGUMENT',
    'Invalid validation range transformation',
  );
  let start = range[startKey],
    end = range[endKey];
  if (operation === 'insert') {
    if (index <= start) {
      start += count;
      end += count;
    } else if (index < end) end += count;
  } else {
    const map = (position: number) =>
      position <= index ? position : position >= index + count ? position - count : index;
    start = map(start);
    end = map(end);
    if (start === end) return null;
  }
  return { ...range, [startKey]: start, [endKey]: end };
}
