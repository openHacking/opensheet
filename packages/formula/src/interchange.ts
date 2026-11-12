import { functionDefinition } from './functions.js';
import { parseFormula, parseSpreadsheetFormula } from './parser.js';
import { printFormulaExplicit } from './printer.js';
import type { AST } from './types.js';

/** Parenthesize supported expressions while preserving one precedence profile in the other. */
export function normalizeFormulaPrecedence(
  expression: string,
  source: 'native' | 'spreadsheet',
): string | undefined {
  if (typeof expression !== 'string' || (source !== 'native' && source !== 'spreadsheet'))
    return undefined;
  try {
    const tree = (source === 'native' ? parseFormula : parseSpreadsheetFormula)(expression);
    const nodes: AST[] = [tree];
    while (nodes.length) {
      const node = nodes.pop()!;
      // A parsed error literal may have originated as an overflowed number.
      // Unknown function names may require an OOXML prefix we cannot recreate.
      if (node.kind === 'literal' && typeof node.value === 'object' && node.value !== null)
        return undefined;
      if (node.kind === 'call') {
        if (!functionDefinition(node.name)) return undefined;
        nodes.push(...node.args);
      } else if (node.kind === 'unary') nodes.push(node.value);
      else if (node.kind === 'binary') nodes.push(node.left, node.right);
    }
    const result = printFormulaExplicit(tree);
    if (result.length > 32768) return undefined;
    const roundTrip = (source === 'native' ? parseSpreadsheetFormula : parseFormula)(result);
    return JSON.stringify(roundTrip) === JSON.stringify(tree) ? result : undefined;
  } catch {
    // Preserve unsupported input unchanged and let the caller report the risk.
    return undefined;
  }
}
