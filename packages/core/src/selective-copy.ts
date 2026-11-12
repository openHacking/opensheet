import { offsetFormula } from '@opensheetjs/formula';
import { assert, toInput, type CellInput } from './types.js';
import type { CellData } from './workbook.js';

/** Convert one captured source cell without mutating the source or destination. */
export function selectiveCopyCell(
  cell: CellData,
  mode: 'values' | 'formulas' | 'formats',
  rowOffset: number,
  columnOffset: number,
) {
  const position = { row: cell.row + rowOffset, column: cell.column + columnOffset };
  if (mode === 'formats')
    return {
      ...position,
      style: {
        ...cell.style,
        ...(cell.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
      },
    };
  const value = cell.value ?? null;
  assert(
    !(mode === 'values' && value && typeof value === 'object' && value.error === '#LIMIT!'),
    'LIMIT_EXCEEDED',
    'Cannot copy an incomplete calculation',
  );
  let input: CellInput =
    mode === 'values'
      ? value && typeof value === 'object'
        ? { type: 'error', code: value.error }
        : toInput(value)
      : (cell.input ?? { type: 'blank' });
  if (input.type === 'formula')
    input = {
      type: 'formula',
      expression: offsetFormula(input.expression, rowOffset, columnOffset),
    };
  return { ...position, input };
}
