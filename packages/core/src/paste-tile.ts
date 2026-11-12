import { createFormulaOffsetter, offsetFormula } from '@opensheetjs/formula';
import { intersects } from './address.js';
import { checkBounds } from './model.js';
import { assert, LIMITS, type Rect, type SheetSnapshot } from './types.js';
import type { PasteCell } from './commands/helpers.js';

export type PasteTileCell = Omit<PasteCell, 'row' | 'column'> & {
  source?: { row: number; column: number };
};
export type PasteTile = {
  sheetId: string;
  range: Rect;
  rows: number;
  columns: number;
  cells: PasteTileCell[];
};

/** Called after command-schema validation and before any page is staged. */
export function validatePasteTile(sheet: SheetSnapshot, tile: PasteTile): number {
  const { range, rows, columns, cells } = tile;
  checkBounds(sheet, range);
  const height = range.endRow - range.startRow;
  const width = range.endColumn - range.startColumn;
  assert(height * width <= LIMITS.rangeCells, 'LIMIT_EXCEEDED', 'Paste dimensions are too large');
  assert(
    cells.length === rows * columns && height % rows === 0 && width % columns === 0,
    'INVALID_ARGUMENT',
    'Paste tile dimensions do not match',
  );
  assert(
    !sheet.merges.some((merge) => intersects(merge, range)),
    'INVALID_RANGE',
    'Unmerge before pasting cell inputs or formats',
  );
  let units = 0;
  for (const cell of cells) units += JSON.stringify(cell).length + 32;
  const total = units * ((height * width) / cells.length);
  assert(total <= 32 * 1024 * 1024, 'LIMIT_EXCEEDED', 'Expanded clipboard content is too large');
  return total;
}

/** Bound the per-command cache by both source text and number of parsed formulas. */
export function createTileOffsetter() {
  const cache = new Map<string, (rows: number, columns: number) => string>();
  let units = 0;
  return (expression: string, rows: number, columns: number): string => {
    let offset = cache.get(expression);
    if (!offset) {
      if (cache.size >= 128 || units + expression.length > 256 * 1024)
        return offsetFormula(expression, rows, columns);
      offset = createFormulaOffsetter(expression);
      cache.set(expression, offset);
      units += expression.length;
    }
    return offset(rows, columns);
  };
}

export function tileCell(
  tile: PasteTile,
  row: number,
  column: number,
  charge: (extra: number) => void,
  offset: (expression: string, rows: number, columns: number) => string = offsetFormula,
): PasteCell {
  const sourceCell =
    tile.cells[
      ((row - tile.range.startRow) % tile.rows) * tile.columns +
        ((column - tile.range.startColumn) % tile.columns)
    ];
  const { source, ...properties } = sourceCell;
  if (source && properties.input?.type === 'formula') {
    const expression = offset(
      properties.input.expression,
      row - source.row,
      column - source.column,
    );
    assert(
      expression.length > 0 && expression.length <= LIMITS.text,
      'LIMIT_EXCEEDED',
      'Pasted formula exceeds the text limit',
    );
    charge(Math.max(0, expression.length - properties.input.expression.length));
    return { ...properties, row, column, input: { type: 'formula', expression } };
  }
  return { ...properties, row, column };
}
