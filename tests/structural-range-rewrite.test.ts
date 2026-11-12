import { expect, it } from 'vitest';
import { parseFormula } from '../packages/formula/src/parser.js';
import {
  structuralRangeReferences,
  structuralReferences,
} from '../packages/core/src/commands/helpers.js';
import type { Command } from '../packages/core/src/types.js';

it('preserves surviving members of every small forward and reversed range', () => {
  for (const axis of ['row', 'column'] as const) {
    for (let first = 0; first < 8; first++)
      for (let last = first; last < 8; last++) {
        for (let at = 0; at < 8; at++)
          for (let count = 1; count <= 8 - at; count++) {
            const command = {
              type: 'core.axis.delete',
              payload: { axis, index: at, count },
            } as Command;
            const transform = structuralRangeReferences(command, 'Data')!;
            const ref = (position: number) => ({
              kind: 'ref' as const,
              row: axis === 'row' ? position : 0,
              column: axis === 'column' ? position : 0,
              absoluteRow: true,
              absoluteColumn: false,
            });
            const expected = Array.from({ length: last - first + 1 }, (_, i) => i + first)
              .filter((position) => position < at || position >= at + count)
              .map((position) => (position < at ? position : position - count));
            for (const reverse of [false, true]) {
              const result = transform(
                ref(reverse ? last : first),
                ref(reverse ? first : last),
                'Data',
              );
              if (!expected.length) expect(result).toBeNull();
              else {
                expect(result?.map((item) => item[axis])).toEqual(
                  reverse ? [expected.at(-1), expected[0]] : [expected[0], expected.at(-1)],
                );
                expect(result?.every((item) => item.absoluteRow && !item.absoluteColumn)).toBe(
                  true,
                );
              }
            }
          }
      }
  }
});

it('leaves ranges on other sheets alone and keeps scalar deletion behavior', () => {
  const command = {
    type: 'core.axis.delete',
    payload: { axis: 'row', index: 1, count: 2 },
  } as Command;
  const range = parseFormula('Other!A2:A4');
  if (range.kind !== 'range') throw new Error('expected range');
  expect(structuralRangeReferences(command, 'Data')!(range.start, range.end, 'Data')).toEqual([
    range.start,
    range.end,
  ]);
  expect(structuralReferences(command, 'Data')({ ...range.start, sheet: 'Data' })).toBeNull();
});
