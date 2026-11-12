import { expect, it } from 'vitest';
import fc from 'fast-check';
import { transformValidationRange } from '../packages/core/src/validation-range.js';

it('maps insert/delete boundaries and removes fully deleted ranges', () => {
  const range = { startRow: 2, endRow: 5, startColumn: 1, endColumn: 3 };
  expect(transformValidationRange(range, 'row', 'insert', 2, 2, 10)).toEqual({
    ...range,
    startRow: 4,
    endRow: 7,
  });
  expect(transformValidationRange(range, 'row', 'insert', 3, 2, 10)).toEqual({
    ...range,
    endRow: 7,
  });
  expect(transformValidationRange(range, 'row', 'insert', 5, 2, 10)).toEqual(range);
  expect(transformValidationRange(range, 'row', 'delete', 1, 3, 10)).toEqual({
    ...range,
    startRow: 1,
    endRow: 2,
  });
  expect(transformValidationRange(range, 'row', 'delete', 2, 3, 10)).toBeNull();
  expect(transformValidationRange(range, 'column', 'delete', 0, 2, 5)).toEqual({
    ...range,
    startColumn: 0,
    endColumn: 1,
  });
  expect(range).toEqual({ startRow: 2, endRow: 5, startColumn: 1, endColumn: 3 });
  expect(() => transformValidationRange(range, 'row', 'delete', 0, 10, 10)).toThrow();
});

it('matches a seeded independent membership model across axis operations', () => {
  fc.assert(
    fc.property(
      fc.integer({ min: 2, max: 40 }),
      fc.nat(),
      fc.nat(),
      fc.nat(),
      fc.nat(),
      fc.boolean(),
      fc.boolean(),
      (length, a, b, c, d, row, inserting) => {
        const start = a % length,
          end = start + 1 + (b % (length - start));
        const index = c % (inserting ? length + 1 : length);
        const count = inserting ? 1 + (d % 4) : 1 + (d % Math.min(length - index, length - 1));
        const source = Array.from({ length }, (_, i) => i >= start && i < end);
        const expected = source.slice();
        if (inserting)
          expected.splice(index, 0, ...Array(count).fill(index > start && index < end));
        else expected.splice(index, count);
        const range = row
          ? { startRow: start, endRow: end, startColumn: 0, endColumn: 1 }
          : { startRow: 0, endRow: 1, startColumn: start, endColumn: end };
        const actual = transformValidationRange(
          range,
          row ? 'row' : 'column',
          inserting ? 'insert' : 'delete',
          index,
          count,
          length,
        );
        const actualMembership = expected.map(
          (_, i) =>
            actual !== null &&
            i >= (row ? actual.startRow : actual.startColumn) &&
            i < (row ? actual.endRow : actual.endColumn),
        );
        expect(actualMembership).toEqual(expected);
      },
    ),
    { seed: 20261031, numRuns: 2000 },
  );
});
