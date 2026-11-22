import { checkBounds, OpenSheetError, type Rect, type Workbook } from '@opensheetjs/core';

/** Bound each Worker response even when cells contain near-limit text or notes. */
export function* clipboardBatches(book: Workbook, sheetId: string, bounds: Rect): Generator<Rect> {
  checkBounds(book.sheetData(sheetId), bounds);
  const width = bounds.endColumn - bounds.startColumn;
  const height = bounds.endRow - bounds.startRow;
  if (width * height > 16384)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Read ranges in batches of at most 16384 cells');
  const batchCells = 256;
  if (width <= batchCells) {
    const rows = Math.max(1, Math.floor(batchCells / width));
    for (let row = bounds.startRow; row < bounds.endRow; row += rows)
      yield { ...bounds, startRow: row, endRow: Math.min(bounds.endRow, row + rows) };
  } else {
    // Row-major ordering keeps clipboard output stable when a row needs several reads.
    for (let row = bounds.startRow; row < bounds.endRow; row++)
      for (let column = bounds.startColumn; column < bounds.endColumn; column += batchCells)
        yield {
          startRow: row,
          endRow: row + 1,
          startColumn: column,
          endColumn: Math.min(bounds.endColumn, column + batchCells),
        };
  }
}
