import type { Rect } from '@opensheetjs/core';

/** Index non-overlapping merges for the rows actually rendered in this batch. */
export function mergeLookup(merges: readonly Rect[], bounds: Rect, rowCount: number) {
  const rows = new Map<number, Rect[]>();
  const startRow = bounds.startRow;
  const endRow = startRow + rowCount;
  for (const merge of merges) {
    if (merge.endColumn <= bounds.startColumn || merge.startColumn >= bounds.endColumn) continue;
    for (
      let row = Math.max(startRow, merge.startRow);
      row < Math.min(endRow, merge.endRow);
      row++
    ) {
      let entries = rows.get(row);
      if (!entries) rows.set(row, (entries = []));
      entries.push(merge);
    }
  }
  for (const entries of rows.values()) entries.sort((a, b) => a.startColumn - b.startColumn);
  let previousRow = -1;
  let previousColumn = -1;
  let entries: Rect[] = [];
  let cursor = 0;
  // Renderers visit columns in ascending order; tolerate a new row or a backward lookup.
  return (row: number, column: number): Rect | undefined => {
    if (row !== previousRow || column < previousColumn) {
      entries = rows.get(row) ?? [];
      cursor = 0;
    }
    previousRow = row;
    previousColumn = column;
    while (cursor < entries.length && entries[cursor].endColumn <= column) cursor++;
    const entry = entries[cursor];
    return entry && entry.startColumn <= column ? entry : undefined;
  };
}
