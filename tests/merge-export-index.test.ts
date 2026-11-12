import { describe, expect, it } from 'vitest';
import { mergeLookup } from '../packages/formats/src/generators/merge-index.js';

describe('merged export lookup', () => {
  it('matches a linear lookup for unsorted, disjoint rectangles across batches and skipped columns', () => {
    const merges = Array.from({ length: 500 }, (_, i) => {
      const row = Math.floor(i / 10) * 3;
      const column = (i % 10) * 3;
      return { startRow: row, endRow: row + 2, startColumn: column, endColumn: column + 2 };
    }).reverse();
    for (let startRow = 0; startRow < 150; startRow += 17) {
      const bounds = {
        startRow,
        endRow: Math.min(150, startRow + 17),
        startColumn: 4,
        endColumn: 25,
      };
      const lookup = mergeLookup(merges, bounds, bounds.endRow - startRow);
      for (let row = bounds.startRow; row < bounds.endRow; row++)
        for (let column = bounds.startColumn; column < bounds.endColumn; column++)
          expect(lookup(row, column)).toEqual(
            merges.find(
              (m) =>
                row >= m.startRow &&
                row < m.endRow &&
                column >= m.startColumn &&
                column < m.endColumn,
            ),
          );
      // A backward lookup must not reuse the cursor from the previous scan.
      expect(lookup(startRow, 4)).toEqual(
        merges.find(
          (m) =>
            startRow >= m.startRow && startRow < m.endRow && 4 >= m.startColumn && 4 < m.endColumn,
        ),
      );
    }
  });
});
