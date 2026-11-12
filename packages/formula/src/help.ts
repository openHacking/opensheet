import { functionDefinition, functionDefinitions, type FunctionDefinition } from './functions.js';

export interface FormulaHelp {
  completion?: { start: number; end: number; functions: readonly FunctionDefinition[] };
  call?: { definition: FunctionDefinition; argument: number };
}

/** Bounded, non-evaluating help for incomplete formulas; offsets are UTF-16 input positions. */
export function formulaHelp(source: string, cursor = source.length): FormulaHelp {
  if (
    !source.startsWith('=') ||
    source.length > 32767 ||
    !Number.isInteger(cursor) ||
    cursor < 1 ||
    cursor > source.length
  )
    return {};
  const stack: Array<{ definition?: FunctionDefinition; argument: number }> = [];
  let quote = '';
  for (let i = 1; i < cursor; i++) {
    const ch = source[i]!;
    if (quote) {
      if (ch === quote) {
        if (source[i + 1] === quote && i + 1 < cursor) i++;
        else quote = '';
      }
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === '(') {
      let end = i;
      while (end > 1 && /\s/.test(source[end - 1]!)) end--;
      let start = end;
      while (start > 1 && /[A-Za-z0-9_.]/.test(source[start - 1]!)) start--;
      stack.push({
        definition: functionDefinition(source.slice(start, end).toUpperCase()),
        argument: 0,
      });
    } else if (ch === ')') stack.pop();
    else if (ch === ',' && stack.length) stack[stack.length - 1]!.argument++;
  }
  const result: FormulaHelp = {};
  const current = [...stack].reverse().find((entry) => entry.definition);
  if (current?.definition)
    result.call = { definition: current.definition, argument: current.argument };
  if (quote) return result;
  let start = cursor,
    end = cursor;
  while (start > 1 && /[A-Za-z0-9_.]/.test(source[start - 1]!)) start--;
  while (end < source.length && /[A-Za-z0-9_.]/.test(source[end]!)) end++;
  const prefix = source.slice(start, cursor).toUpperCase();
  if (
    prefix &&
    /^[A-Z][A-Z0-9_.]*$/.test(prefix) &&
    /[=,+\-*/^&<>(:]|\s/.test(source[start - 1]!)
  ) {
    const functions = functionDefinitions
      .filter((definition) => definition.name.startsWith(prefix))
      .sort((a, b) => a.name.localeCompare(b.name))
      .slice(0, 8);
    if (functions.length) result.completion = { start, end, functions };
  }
  return result;
}
