/** Deterministic, bounded formula parser. Never evaluates JavaScript. */
export type FormulaValue = string | number | boolean | null | { error: string };
export type Reference = {
  kind: 'ref';
  row: number;
  column: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
  sheet?: string;
};
export type AST =
  | Reference
  | { kind: 'literal'; value: FormulaValue }
  | { kind: 'range'; start: Reference; end: Reference }
  | { kind: 'unary'; op: string; value: AST }
  | { kind: 'binary'; op: string; left: AST; right: AST }
  | { kind: 'call'; name: string; args: AST[] };
export class FormulaError extends Error {
  constructor(public code: string) {
    super(code);
  }
}
const fail = (code = '#VALUE!'): never => {
  throw new FormulaError(code);
};
export function columnName(index: number): string {
  let s = '';
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26))
    s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
export function columnIndex(name: string): number {
  return [...name.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
}
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
function checked(v: FormulaValue): FormulaValue {
  if (v !== null && typeof v === 'object') fail(v.error);
  return v;
}
function number(v: FormulaValue): number {
  checked(v);
  if (v === null || v === '') return 0;
  if (typeof v === 'boolean') return +v;
  const n = Number(v);
  if (!Number.isFinite(n)) fail();
  return n;
}
function truth(v: FormulaValue): boolean {
  checked(v);
  if (typeof v === 'string') {
    if (v.toUpperCase() === 'TRUE') return true;
    if (v.toUpperCase() === 'FALSE') return false;
    fail();
  }
  return !!v;
}
export const supportedFunctions = [
  'SUM',
  'AVERAGE',
  'MIN',
  'MAX',
  'COUNT',
  'COUNTA',
  'IF',
  'IFERROR',
  'AND',
  'OR',
  'NOT',
  'ROUND',
  'ABS',
] as const;
export function evaluateFormula(
  ast: AST,
  read: (ref: Reference) => FormulaValue,
  maxCells = 200000,
): FormulaValue {
  let budget = maxCells;
  type Result = FormulaValue | FormulaValue[];
  const single = (v: Result): FormulaValue => (Array.isArray(v) ? fail() : v);
  function run(n: AST): Result {
    if (--budget < 0) fail('#LIMIT!');
    if (n.kind === 'literal') return n.value;
    if (n.kind === 'ref') return read(n);
    if (n.kind === 'range') {
      const out: FormulaValue[] = [];
      for (let r = Math.min(n.start.row, n.end.row); r <= Math.max(n.start.row, n.end.row); r++)
        for (
          let c = Math.min(n.start.column, n.end.column);
          c <= Math.max(n.start.column, n.end.column);
          c++
        ) {
          if (--budget < 0) fail('#LIMIT!');
          out.push(read({ ...n.start, row: r, column: c }));
        }
      return out;
    }
    if (n.kind === 'unary') {
      const v = number(single(run(n.value)));
      return n.op === '-' ? -v : n.op === '%' ? v / 100 : v;
    }
    if (n.kind === 'binary') {
      const a = checked(single(run(n.left))),
        b = checked(single(run(n.right)));
      if (n.op === '&') return String(a ?? '') + String(b ?? '');
      if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
        const x = typeof a === 'string' ? a.toLowerCase() : (a ?? 0),
          y = typeof b === 'string' ? b.toLowerCase() : (b ?? 0);
        const cmp =
          x === y ? 0 : typeof x === typeof y ? (x < y ? -1 : 1) : typeof x === 'number' ? -1 : 1;
        return n.op === '='
          ? cmp === 0
          : n.op === '<>'
            ? cmp !== 0
            : n.op === '<'
              ? cmp < 0
              : n.op === '>'
                ? cmp > 0
                : n.op === '<='
                  ? cmp <= 0
                  : cmp >= 0;
      }
      const x = number(a),
        y = number(b);
      const result =
        n.op === '+'
          ? x + y
          : n.op === '-'
            ? x - y
            : n.op === '*'
              ? x * y
              : n.op === '/'
                ? y === 0
                  ? fail('#DIV/0!')
                  : x / y
                : x ** y;
      return Number.isFinite(result) ? result : fail('#NUM!');
    }
    const arity = (min: number, max = min) => {
      if (n.args.length < min || n.args.length > max) fail();
    };
    if (n.name === 'IF') {
      arity(2, 3);
      return truth(single(run(n.args[0]))) ? run(n.args[1]) : n.args[2] ? run(n.args[2]) : false;
    }
    if (n.name === 'IFERROR') {
      arity(2);
      try {
        const v = run(n.args[0]);
        checked(single(v));
        return v;
      } catch (e) {
        if (e instanceof FormulaError) return run(n.args[1]);
        throw e;
      }
    }
    if (!(supportedFunctions as readonly string[]).includes(n.name)) fail('#NAME?');
    if (n.name === 'NOT' || n.name === 'ABS') arity(1);
    else if (n.name === 'ROUND') arity(2);
    else arity(1, 255);
    const vals = n.args.flatMap((a) => {
      const value = run(a);
      return Array.isArray(value) ? value : [value];
    });
    if (n.name === 'COUNTA') return vals.filter((v) => v !== null).length;
    if (n.name === 'COUNT') return vals.filter((v) => typeof v === 'number').length;
    vals.forEach(checked);
    if (n.name === 'AND') return vals.every(truth);
    if (n.name === 'OR') return vals.some(truth);
    if (n.name === 'NOT') return !truth(vals[0]);
    if (n.name === 'ABS') return Math.abs(number(vals[0]));
    if (n.name === 'ROUND') {
      const digits = number(vals[1]);
      if (!Number.isInteger(digits) || Math.abs(digits) > 100) fail('#NUM!');
      const scale = 10 ** digits,
        v = number(vals[0]);
      const result = (Math.sign(v) * Math.round(Math.abs(v) * scale + Number.EPSILON)) / scale;
      return Number.isFinite(result) ? result : fail('#NUM!');
    }
    const nums = vals.filter((v): v is number => typeof v === 'number');
    if (n.name === 'SUM') return nums.reduce((a, b) => a + b, 0);
    if (n.name === 'AVERAGE')
      return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : fail('#DIV/0!');
    if (n.name === 'MIN') return nums.length ? nums.reduce((a, b) => Math.min(a, b)) : 0;
    if (n.name === 'MAX') return nums.length ? nums.reduce((a, b) => Math.max(a, b)) : 0;
    return fail('#NAME?');
  }
  try {
    const result = single(run(ast));
    return typeof result === 'number' && !Number.isFinite(result) ? { error: '#NUM!' } : result;
  } catch (e) {
    if (e instanceof FormulaError) return { error: e.code };
    throw e;
  }
}
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
