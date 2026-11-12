import { fail, FormulaError } from './errors.js';
import type { AST, FormulaValue, Reference } from './types.js';

import { checked, number, truth } from './coercion.js';
import { functionDefinition } from './functions.js';
import { Aggregate } from './aggregate.js';
import { lookup } from './lookup.js';
import { conditional } from './conditional.js';
import { evaluateDate, type DateSystem } from './dates.js';
import { directedRound, modulo, nearestRound } from './numeric.js';
import { evaluateText, formulaText, FORMULA_TEXT_LIMIT, TextJoin } from './text.js';
export { supportedFunctions } from './functions.js';

export function evaluateFormula(
  ast: AST,
  read: (ref: Reference) => FormulaValue,
  maxCells = 200000,
  dateSystem: DateSystem = '1900',
): FormulaValue {
  let budget = maxCells;
  let textWork: { remaining: number } | undefined;
  type Result = FormulaValue | FormulaValue[];
  const single = (v: Result): FormulaValue => (Array.isArray(v) ? fail() : v);
  const readReference = (ref: Reference): FormulaValue => {
    const value = read(ref);
    return typeof value === 'number' && !Number.isFinite(value) ? { error: '#NUM!' } : value;
  };
  function run(n: AST): Result {
    if (--budget < 0) fail('#LIMIT!');
    if (n.kind === 'literal' || n.kind === 'ref') {
      const value = n.kind === 'literal' ? n.value : readReference(n);
      return typeof value === 'number' && !Number.isFinite(value) ? fail('#NUM!') : value;
    }
    if (n.kind === 'range') return fail();
    if (n.kind === 'unary') {
      const v = number(single(run(n.value)));
      return n.op === '-' ? -v : n.op === '%' ? v / 100 : v;
    }
    if (n.kind === 'binary') {
      const a = checked(single(run(n.left))),
        b = checked(single(run(n.right)));
      if (n.op === '&') {
        const result = formulaText(a) + formulaText(b);
        return result.length <= FORMULA_TEXT_LIMIT ? result : fail();
      }
      if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
        const x = typeof a === 'string' ? a.toLowerCase() : (a ?? 0),
          y = typeof b === 'string' ? b.toLowerCase() : (b ?? 0);
        const rank = (value: typeof x) =>
          typeof value === 'number' ? 0 : typeof value === 'string' ? 1 : 2;
        const cmp = x === y ? 0 : typeof x === typeof y ? (x < y ? -1 : 1) : rank(x) - rank(y);
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
    const definition = functionDefinition(n.name);
    if (!definition) return fail('#NAME?');
    if (n.args.length < definition.minArgs || n.args.length > definition.maxArgs) fail();
    if (definition.evaluation === 'conditional' || definition.evaluation === 'lookup') {
      const iterator = (definition.evaluation === 'lookup' ? lookup : conditional)(
        n.name,
        n.args,
        (textWork ??= { remaining: 1_000_000 }),
      );
      let step = iterator.next();
      while (!step.done) step = iterator.next(single(run(step.value)));
      return step.value;
    }
    if (n.name === 'TRUE' || n.name === 'FALSE') return n.name === 'TRUE';
    if (n.name === 'IF') {
      return single(
        truth(single(run(n.args[0]))) ? run(n.args[1]) : n.args[2] ? run(n.args[2]) : false,
      );
    }
    if (n.name === 'IFS') {
      if (n.args.length % 2) fail();
      for (let i = 0; i < n.args.length; i += 2)
        if (truth(single(run(n.args[i])))) return single(run(n.args[i + 1]));
      return fail('#N/A');
    }
    if (n.name === 'IFERROR' || n.name === 'IFNA') {
      try {
        const v = run(n.args[0]);
        checked(single(v));
        return v;
      } catch (e) {
        if (
          e instanceof FormulaError &&
          e.code !== '#LIMIT!' &&
          e.code !== '#CYCLE!' &&
          (n.name === 'IFERROR' || e.code === '#N/A')
        )
          return run(n.args[1]);
        throw e;
      }
    }
    if (definition.evaluation === 'aggregate' || definition.evaluation === 'text') {
      const argument = (node: AST) => {
        try {
          return single(run(node));
        } catch (error) {
          if (!(error instanceof FormulaError) || error.code === '#LIMIT!') throw error;
          return { error: error.code };
        }
      };
      const aggregate =
        definition.evaluation === 'aggregate'
          ? new Aggregate(n.name)
          : n.name === 'TEXTJOIN'
            ? new TextJoin(argument(n.args[0]), argument(n.args[1]))
            : new TextJoin();
      for (const arg of n.name === 'TEXTJOIN' ? n.args.slice(2) : n.args) {
        if (arg.kind === 'range') {
          if (--budget < 0) fail('#LIMIT!');
          for (
            let row = Math.min(arg.start.row, arg.end.row);
            row <= Math.max(arg.start.row, arg.end.row);
            row++
          )
            for (
              let column = Math.min(arg.start.column, arg.end.column);
              column <= Math.max(arg.start.column, arg.end.column);
              column++
            ) {
              if (--budget < 0) fail('#LIMIT!');
              aggregate.add(readReference({ ...arg.start, row, column }), true);
            }
          continue;
        }
        let value: Result;
        try {
          value = run(arg);
        } catch (error) {
          if (!(error instanceof FormulaError) || error.code === '#LIMIT!') throw error;
          value = { error: error.code };
        }
        aggregate.add(single(value), arg.kind === 'ref');
      }
      return aggregate.result();
    }
    const vals = n.args.map((arg) => checked(single(run(arg))));
    const date = evaluateDate(n.name, vals, dateSystem);
    if (date !== undefined) return date;
    const text = evaluateText(n.name, vals);
    if (text !== undefined) return text;
    if (n.name === 'ROUNDUP' || n.name === 'ROUNDDOWN')
      return directedRound(number(vals[0]), number(vals[1]), n.name === 'ROUNDUP');
    if (n.name === 'INT') return Math.floor(number(vals[0])) || 0;
    if (n.name === 'MOD') return modulo(number(vals[0]), number(vals[1]));
    if (n.name === 'NOT') return !truth(vals[0]);
    if (n.name === 'ABS') return Math.abs(number(vals[0]));
    if (n.name === 'ROUND') {
      const digits = Math.trunc(number(vals[1]));
      return nearestRound(number(vals[0]), digits);
    }
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
