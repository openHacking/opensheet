import { checked } from './coercion.js';
import { criterion } from './criteria.js';
import { fail } from './errors.js';
import { area, cell, type Area } from './ranges.js';
import type { AST, FormulaValue } from './types.js';

export function conditionalPlan(name: string, args: AST[]) {
  const count = name.startsWith('COUNT');
  const multiple = name.endsWith('IFS');
  const ranges: Area[] = [],
    tests: AST[] = [];
  let values: Area | undefined;
  if (!multiple) {
    ranges.push(area(args[0]));
    tests.push(args[1]);
    if (!count) {
      values = args[2] ? area(args[2]) : ranges[0];
      // SUMIF/AVERAGEIF extend the value range from its top-left to the criteria shape.
      values = { ...values, rows: ranges[0].rows, columns: ranges[0].columns };
    }
  } else {
    const offset = count ? 0 : 1;
    if ((args.length - offset) % 2) fail();
    if (!count) values = area(args[0]);
    for (let i = offset; i < args.length; i += 2) {
      ranges.push(area(args[i]));
      tests.push(args[i + 1]);
    }
    const shape = values ?? ranges[0];
    if (ranges.some((r) => r.rows !== shape.rows || r.columns !== shape.columns)) fail();
  }
  return { ranges, tests, values, count, average: name.startsWith('AVERAGE') };
}
/** Shared synchronous/asynchronous execution plan: only one cell result is retained. */
export function* conditional(
  name: string,
  args: AST[],
  work = { remaining: 1_000_000 },
): Generator<AST, FormulaValue, FormulaValue> {
  const plan = conditionalPlan(name, args);
  const predicates = [];
  for (const test of plan.tests) predicates.push(criterion(yield test, work));
  let sum = 0,
    count = 0;
  const shape = plan.ranges[0];
  for (let row = 0; row < shape.rows; row++)
    for (let column = 0; column < shape.columns; column++) {
      let matches = true;
      for (let i = 0; i < plan.ranges.length; i++) {
        const value = yield cell(plan.ranges[i], row, column);
        if (!predicates[i](value)) {
          matches = false;
          break;
        }
      }
      if (!matches) continue;
      if (plan.count) {
        count++;
        continue;
      }
      const value = checked(yield cell(plan.values!, row, column));
      if (typeof value === 'number') {
        sum += value;
        count++;
      }
    }
  if (plan.count) return count;
  if (plan.average && !count) return { error: '#DIV/0!' };
  const result = plan.average ? sum / count : sum;
  return Number.isFinite(result) ? result : { error: '#NUM!' };
}
