import { columnIndex } from './address.js';
import { fail } from './errors.js';
import type { AST, Reference } from './types.js';

const tokenRE =
  /\s*("(?:[^"]|"")*"|'(?:[^']|'')*'|#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|CYCLE!|LIMIT!)|\$?[A-Za-z]+\$?[1-9]\d*|(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[A-Za-z_][A-Za-z_0-9.]*|<=|>=|<>|[+\-*/^&=<>%(),:!])/gy;
export function parseFormula(expression: string): AST {
  if (expression.length > 32768) fail('#LIMIT!');
  const text = expression.replace(/^=/, '').trim();
  const tokens: string[] = [];
  let pos = 0;
  while (pos < text.length) {
    tokenRE.lastIndex = pos;
    const m = tokenRE.exec(text);
    if (!m) fail();
    tokens.push(m![1]);
    pos = tokenRE.lastIndex;
    if (tokens.length > 4096) fail('#LIMIT!');
  }
  let i = 0,
    depth = 0;
  const peek = () => tokens[i];
  const next = () => tokens[i++];
  const reference = (t: string, sheet?: string): Reference => {
    const m = /^(\$?)([A-Za-z]+)(\$?)([1-9]\d*)$/.exec(t);
    if (!m) fail('#REF!');
    const row = Number(m![4]) - 1,
      column = columnIndex(m![2]);
    if (row >= 1048576 || column >= 16384) fail('#REF!');
    return {
      kind: 'ref',
      row,
      column,
      absoluteRow: !!m![3],
      absoluteColumn: !!m![1],
      ...(sheet ? { sheet } : {}),
    };
  };
  const precedence: Record<string, number> = {
    '=': 1,
    '<>': 1,
    '<': 1,
    '>': 1,
    '<=': 1,
    '>=': 1,
    '&': 2,
    '+': 3,
    '-': 3,
    '*': 4,
    '/': 4,
    '^': 5,
  };
  function atom(): AST {
    if (++depth > 128) fail('#LIMIT!');
    const t = next();
    if (!t) fail();
    let n: AST;
    if (t === '(') {
      n = expr(0);
      if (next() !== ')') fail();
    } else if (t === '+' || t === '-') n = { kind: 'unary', op: t, value: expr(5) };
    else if (t.startsWith('"')) n = { kind: 'literal', value: t.slice(1, -1).replace(/""/g, '"') };
    else if (t.startsWith('#')) n = { kind: 'literal', value: { error: t } };
    else if (peek() === '!') {
      next();
      const sheet = t.startsWith("'") ? t.slice(1, -1).replace(/''/g, "'") : t;
      n = reference(next(), sheet);
    } else if (peek() === '(') {
      next();
      const args: AST[] = [];
      if (peek() !== ')') {
        do {
          args.push(expr(0));
          if (peek() !== ',') break;
          next();
        } while (true);
      }
      if (next() !== ')') fail();
      n = { kind: 'call', name: t.toUpperCase(), args };
    } else if (/^\$?[A-Za-z]+\$?\d+$/.test(t)) n = reference(t);
    else if (/^(TRUE|FALSE)$/i.test(t)) n = { kind: 'literal', value: t.toUpperCase() === 'TRUE' };
    else if (/^\d|^\.\d/.test(t)) n = { kind: 'literal', value: Number(t) };
    else fail('#NAME?');
    if (n!.kind === 'ref' && peek() === ':') {
      next();
      const end = reference(next(), n!.sheet);
      n = { kind: 'range', start: n!, end };
    }
    while (peek() === '%') {
      next();
      n = { kind: 'unary', op: '%', value: n! };
    }
    depth--;
    return n!;
  }
  function expr(min: number): AST {
    let left = atom();
    while (peek() && precedence[peek()] >= min) {
      const op = next(),
        p = precedence[op];
      left = { kind: 'binary', op, left, right: expr(p + (op === '^' ? 0 : 1)) };
    }
    return left;
  }
  const ast = expr(0);
  if (i !== tokens.length) fail();
  return ast;
}
