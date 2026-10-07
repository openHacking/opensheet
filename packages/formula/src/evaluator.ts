import { fail, FormulaError } from './errors.js';
import type { AST, FormulaValue, Reference } from './types.js';

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
