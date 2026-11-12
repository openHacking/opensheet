import type { AST } from '@opensheetjs/formula';

type Reference = Extract<AST, { kind: 'ref' }>;

/** Transform an inclusive formula range after deleting one axis span. */
export function transformDeletedFormulaRange(
  start: Reference,
  end: Reference,
  axis: 'row' | 'column',
  index: number,
  count: number,
): [Reference, Reference] | null {
  const lo = Math.min(start[axis], end[axis]);
  const hi = Math.max(start[axis], end[axis]);
  const after = index + count;
  if (lo >= index && hi < after) return null;
  const first = lo < index ? lo : Math.max(lo, after) - count;
  const last = hi >= after ? hi - count : Math.min(hi, index - 1);
  const forward = start[axis] <= end[axis];
  return [
    { ...start, [axis]: forward ? first : last },
    { ...end, [axis]: forward ? last : first },
  ];
}
