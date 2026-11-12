import { expect, it, vi } from 'vitest';
import { validationCandidates } from '../packages/core/src/validation-candidates.js';

const rule = {
  kind: 'number' as const,
  operator: 'gte' as const,
  first: 1,
  allowBlank: false,
  policy: 'reject' as const,
};

it('collects the deduplicated intersection of edited ranges and rules', async () => {
  let seed = 12345;
  const random = (limit: number) => (seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) % limit;
  for (let run = 0; run < 200; run++) {
    const rect = () => {
      const startRow = random(8),
        startColumn = random(7);
      return {
        startRow,
        startColumn,
        endRow: startRow + 1 + random(8 - startRow),
        endColumn: startColumn + 1 + random(7 - startColumn),
      };
    };
    const rules = Array.from({ length: 1 + random(5) }, (_, i) => ({
      id: String(i),
      range: rect(),
      rule,
    }));
    const ranges = Array.from({ length: 1 + random(5) }, rect);
    const expected = [];
    for (let row = 0; row < 8; row++)
      for (let column = 0; column < 7; column++) {
        const contains = (r: ReturnType<typeof rect>) =>
          row >= r.startRow && row < r.endRow && column >= r.startColumn && column < r.endColumn;
        if (ranges.some(contains) && rules.some(({ range }) => contains(range)))
          expected.push({ row, column });
      }
    const actual = await validationCandidates(rules, ranges, { rows: 8, columns: 7 });
    expect(actual.positions).toEqual(expected);
  }
});

it('rejects oversized, invalid and cancelled inputs without mutating callers', async () => {
  const rules = [
    { id: 'one', range: { startRow: 0, endRow: 1000, startColumn: 0, endColumn: 100 }, rule },
  ];
  const original = structuredClone(rules);
  await expect(
    validationCandidates(rules, [rules[0].range], { rows: 1000, columns: 100 }),
  ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  await expect(
    validationCandidates(rules, [{ startRow: 0, endRow: 1001, startColumn: 0, endColumn: 1 }], {
      rows: 1000,
      columns: 100,
    }),
  ).rejects.toMatchObject({ code: 'INVALID_RANGE' });
  const controller = new AbortController();
  controller.abort(new Error('cancelled'));
  await expect(
    validationCandidates(rules, [], { rows: 1000, columns: 100 }, controller.signal),
  ).rejects.toThrow('cancelled');
  expect(rules).toEqual(original);
});

it('yields to cancellation while comparing many disjoint ranges', async () => {
  const rules = Array.from({ length: 512 }, (_, i) => ({
    id: String(i),
    range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
    rule,
  }));
  const ranges = Array.from({ length: 1800 }, () => ({
    startRow: 1,
    endRow: 2,
    startColumn: 0,
    endColumn: 1,
  }));
  const controller = new AbortController();
  let clock = 0;
  const now = vi.spyOn(performance, 'now').mockImplementation(() => (clock += 10));
  try {
    setTimeout(() => controller.abort(new Error('cancelled during comparison')), 0);
    await expect(
      validationCandidates(rules, ranges, { rows: 2, columns: 1 }, controller.signal),
    ).rejects.toThrow('cancelled during comparison');
  } finally {
    now.mockRestore();
  }
});
