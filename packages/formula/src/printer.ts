import { columnName } from './address.js';
import type { AST, Reference } from './types.js';

export function printFormula(n: AST): string {
  return print(n, false);
}

/** Parenthesize unary operators as whole operands across precedence profiles. */
export function printFormulaExplicit(n: AST): string {
  return print(n, true);
}

function print(n: AST, explicitUnary: boolean): string {
  const ref = (r: Reference) =>
    `${r.sheet ? `'${r.sheet.replace(/'/g, "''")}'!` : ''}${r.absoluteColumn ? '$' : ''}${columnName(r.column)}${r.absoluteRow ? '$' : ''}${r.row + 1}`;
  switch (n.kind) {
    case 'ref':
      return ref(n);
    case 'range':
      return ref(n.start) + ':' + ref({ ...n.end, sheet: undefined });
    case 'literal':
      return n.value === null
        ? '0'
        : typeof n.value === 'object'
          ? n.value.error
          : typeof n.value === 'string'
            ? `"${n.value.replace(/"/g, '""')}"`
            : String(n.value).toUpperCase();
    case 'unary':
      if (n.op === '%') return `(${print(n.value, explicitUnary)})%`;
      return explicitUnary
        ? `(${n.op}(${print(n.value, explicitUnary)}))`
        : `${n.op}(${print(n.value, explicitUnary)})`;
    case 'binary':
      return `(${print(n.left, explicitUnary)}${n.op}${print(n.right, explicitUnary)})`;
    case 'call':
      return `${n.name}(${n.args.map((arg) => print(arg, explicitUnary)).join(',')})`;
  }
}
