import { fail } from './errors.js';
import type { AST, Reference } from './types.js';
import { FORMULA_COLUMNS, FORMULA_ROWS } from './reference-bounds.js';

export type Area = { start: Reference; rows: number; columns: number };
export function area(ast: AST): Area {
  if (ast.kind === 'ref') return { start: ast, rows: 1, columns: 1 };
  if (ast.kind !== 'range') return fail();
  if (ast.start.sheet !== ast.end.sheet) return fail('#REF!');
  return {
    start: {
      ...ast.start,
      row: Math.min(ast.start.row, ast.end.row),
      column: Math.min(ast.start.column, ast.end.column),
    },
    rows: Math.abs(ast.start.row - ast.end.row) + 1,
    columns: Math.abs(ast.start.column - ast.end.column) + 1,
  };
}
export const cell = (range: Area, row: number, column: number): AST => {
  const targetRow = range.start.row + row;
  const targetColumn = range.start.column + column;
  if (
    targetRow < 0 ||
    targetRow >= FORMULA_ROWS ||
    targetColumn < 0 ||
    targetColumn >= FORMULA_COLUMNS
  )
    return { kind: 'literal', value: { error: '#REF!' } };
  return { ...range.start, row: targetRow, column: targetColumn };
};
