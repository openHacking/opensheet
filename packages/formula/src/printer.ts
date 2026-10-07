import { columnName } from './address.js';
import type { AST, Reference } from './types.js';

export function printFormula(n: AST): string {
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
      return n.op === '%' ? `(${printFormula(n.value)})%` : `${n.op}(${printFormula(n.value)})`;
    case 'binary':
      return `(${printFormula(n.left)}${n.op}${printFormula(n.right)})`;
    case 'call':
      return `${n.name}(${n.args.map(printFormula).join(',')})`;
  }
}
