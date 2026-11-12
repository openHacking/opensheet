import { expect, it } from 'vitest';
import fc from 'fast-check';
import { GridLayout } from '../packages/renderer/src/layout.js';

it('resolves hidden runs and edges like a bounded linear reference', () => {
  const layout = new GridLayout();
  fc.assert(
    fc.property(
      fc.array(fc.boolean(), { minLength: 1, maxLength: 100 }),
      fc.nat(),
      fc.boolean(),
      (hidden, raw, positive) => {
        const offsets = [0];
        for (const item of hidden) offsets.push(offsets.at(-1)! + (item ? 0 : 30));
        const index = raw % hidden.length;
        const direction = positive ? 1 : -1;
        const find = (step: number) => {
          for (let i = index; i >= 0 && i < hidden.length; i += step) if (!hidden[i]) return i;
          return undefined;
        };
        const expected = find(direction) ?? find(-direction) ?? index;
        expect(layout.visibleIndex(offsets, index, direction)).toBe(expected);
      },
    ),
    { seed: 20261112, numRuns: 2000 },
  );
});

it('handles fully hidden and long hidden edges without leaving the axis', () => {
  const layout = new GridLayout();
  const hidden = Array(100001).fill(0);
  expect(layout.visibleIndex(hidden, 0, -1)).toBe(0);
  expect(layout.visibleIndex(hidden, 99999, 1)).toBe(99999);
  hidden[100000] = 30;
  expect(layout.visibleIndex(hidden, 0, -1)).toBe(99999);
});

it('enumerates visible frozen and scrolling entries like the linear painter', () => {
  const layout = new GridLayout();
  fc.assert(
    fc.property(
      fc.array(fc.integer({ min: 0, max: 5 }), { minLength: 1, maxLength: 200 }),
      fc.nat(500),
      fc.nat(500),
      fc.nat(),
      (sizes, scroll, extent, rawFrozen) => {
        const offsets = [0];
        for (const size of sizes) offsets.push(offsets.at(-1)! + size);
        const frozen = rawFrozen % (sizes.length + 1);
        const expected: number[] = [];
        for (let i = 0; i < frozen && offsets[i] < extent; i++)
          if (offsets[i + 1] > offsets[i]) expected.push(i);
        const start = Math.max(frozen, layout.index(offsets, scroll + offsets[frozen]));
        for (let i = start; i < sizes.length && offsets[i] - scroll < extent; i++)
          if (offsets[i + 1] > offsets[i]) expected.push(i);
        expect(layout.visiblePositions(offsets, scroll, extent, frozen)).toEqual(expected);
      },
    ),
    { seed: 20261112, numRuns: 3000 },
  );
});

it('skips a long frozen hidden run with logarithmic offset reads', () => {
  const layout = new GridLayout();
  const offsets = Array(100001).fill(0);
  offsets.push(30);
  let reads = 0;
  const observed = new Proxy(offsets, {
    get(target, property, receiver) {
      if (typeof property === 'string' && /^\d+$/.test(property)) reads++;
      return Reflect.get(target, property, receiver);
    },
  });
  expect(layout.visiblePositions(observed, 0, 100, 100000)).toEqual([100000]);
  expect(reads).toBeLessThan(150);
});
