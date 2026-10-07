import { parseFormula } from './parser.js';
import { printFormula } from './printer.js';
import type { AST, Reference } from './types.js';

export function transformReferences(ast: AST, fn: (ref: Reference) => Reference | null): AST {
  if (ast.kind === 'ref') return fn(ast) ?? { kind: 'literal', value: { error: '#REF!' } };
  if (ast.kind === 'range') {
    const start = fn(ast.start),
      end = fn(ast.end);
    return start && end
      ? { kind: 'range', start, end }
      : { kind: 'literal', value: { error: '#REF!' } };
  }
  if (ast.kind === 'unary') return { ...ast, value: transformReferences(ast.value, fn) };
  if (ast.kind === 'binary')
    return {
      ...ast,
      left: transformReferences(ast.left, fn),
      right: transformReferences(ast.right, fn),
    };
  if (ast.kind === 'call') return { ...ast, args: ast.args.map((a) => transformReferences(a, fn)) };
  return ast;
}
export function offsetFormula(expression: string, rows: number, columns: number): string {
  return printFormula(
    transformReferences(parseFormula(expression), (r) => {
      const row = r.row + (r.absoluteRow ? 0 : rows),
        column = r.column + (r.absoluteColumn ? 0 : columns);
      return row < 0 || column < 0 ? null : { ...r, row, column };
    }),
  );
}
