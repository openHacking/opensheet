import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import { createSnapshot, validateSnapshot } from '../packages/core/src/model.js';

describe('snapshot merge validation', () => {
  it('matches an independent pairwise overlap check for unordered rectangles', () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 12, columns: 10 }] });
    const rectangle = fc
      .record({
        row: fc.integer({ min: 0, max: 11 }),
        column: fc.integer({ min: 0, max: 9 }),
        height: fc.integer({ min: 1, max: 12 }),
        width: fc.integer({ min: 1, max: 10 }),
      })
      .map(({ row, column, height, width }) => ({
        startRow: row,
        endRow: Math.min(12, row + height),
        startColumn: column,
        endColumn: Math.min(10, column + width),
      }));
    fc.assert(
      fc.property(fc.array(rectangle, { maxLength: 30 }), (merges) => {
        const overlap = merges.some((a, i) =>
          merges
            .slice(0, i)
            .some(
              (b) =>
                a.startRow < b.endRow &&
                b.startRow < a.endRow &&
                a.startColumn < b.endColumn &&
                b.startColumn < a.endColumn,
            ),
        );
        const input = structuredClone(snapshot);
        input.sheets[0].merges = merges;
        if (overlap) expect(() => validateSnapshot(input)).toThrow('Overlapping merges');
        else expect(validateSnapshot(input).sheets[0].merges).toEqual(merges);
      }),
      { seed: 20261111, numRuns: 500 },
    );
  });

  it('handles the maximum number of touching, unsorted merges and rejects an overlapping duplicate', () => {
    const input = createSnapshot({ sheets: [{ name: 'Data', rows: 10000, columns: 2 }] });
    input.sheets[0].merges = Array.from({ length: 10000 }, (_, row) => ({
      startRow: row,
      endRow: row + 1,
      startColumn: 0,
      endColumn: 2,
    })).reverse();
    expect(validateSnapshot(input).sheets[0].merges).toEqual(input.sheets[0].merges);
    input.sheets[0].merges[9999] = input.sheets[0].merges[0];
    expect(() => validateSnapshot(input)).toThrow('Overlapping merges');
  });
});
