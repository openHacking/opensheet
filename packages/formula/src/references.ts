import { parseFormula } from './parser.js';
import { printFormula } from './printer.js';
import type { AST, Reference } from './types.js';
import { FORMULA_COLUMNS, FORMULA_ROWS } from './reference-bounds.js';

function checkedReference(ref: Reference, fn: (ref: Reference) => Reference | null) {
  const result = fn(ref);
  return result &&
    Number.isInteger(result.row) &&
    Number.isInteger(result.column) &&
    result.row >= 0 &&
    result.row < FORMULA_ROWS &&
    result.column >= 0 &&
    result.column < FORMULA_COLUMNS
    ? result
    : null;
}

export function transformReferences(
  ast: AST,
  fn: (ref: Reference) => Reference | null,
  transformRange?: (start: Reference, end: Reference) => [Reference, Reference] | null,
): AST {
  if (ast.kind === 'ref')
    return checkedReference(ast, fn) ?? { kind: 'literal', value: { error: '#REF!' } };
  if (ast.kind === 'range') {
    const range = transformRange?.(ast.start, ast.end);
    if (transformRange && !range) return { kind: 'literal', value: { error: '#REF!' } };
    const start = checkedReference(range?.[0] ?? ast.start, transformRange ? (ref) => ref : fn),
      end = checkedReference(range?.[1] ?? ast.end, transformRange ? (ref) => ref : fn);
    return start && end
      ? { kind: 'range', start, end }
      : { kind: 'literal', value: { error: '#REF!' } };
  }
  if (ast.kind === 'unary')
    return { ...ast, value: transformReferences(ast.value, fn, transformRange) };
  if (ast.kind === 'binary')
    return {
      ...ast,
      left: transformReferences(ast.left, fn, transformRange),
      right: transformReferences(ast.right, fn, transformRange),
    };
  if (ast.kind === 'call')
    return { ...ast, args: ast.args.map((a) => transformReferences(a, fn, transformRange)) };
  return ast;
}
function offsetAST(ast: AST, rows: number, columns: number): string {
  return printFormula(
    transformReferences(ast, (r) => {
      const row = r.row + (r.absoluteRow ? 0 : rows),
        column = r.column + (r.absoluteColumn ? 0 : columns);
      return row < 0 || column < 0 ? null : { ...r, row, column };
    }),
  );
}
export function offsetFormula(expression: string, rows: number, columns: number): string {
  return offsetAST(parseFormula(expression), rows, columns);
}

/** Parse once for repeated reference relocation; the returned function owns its immutable AST. */
export function createFormulaOffsetter(
  expression: string,
): (rows: number, columns: number) => string {
  const ast = parseFormula(expression);
  const movable = (node: AST): boolean => {
    if (node.kind === 'ref') return !node.absoluteRow || !node.absoluteColumn;
    if (node.kind === 'range') return movable(node.start) || movable(node.end);
    if (node.kind === 'unary') return movable(node.value);
    if (node.kind === 'binary') return movable(node.left) || movable(node.right);
    if (node.kind === 'call') return node.args.some(movable);
    return false;
  };
  if (!movable(ast)) {
    const result = offsetAST(ast, 0, 0);
    return () => result;
  }
  return (rows, columns) => offsetAST(ast, rows, columns);
}
