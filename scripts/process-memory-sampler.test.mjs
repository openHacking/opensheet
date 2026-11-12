import { test } from 'node:test';
import { deepStrictEqual } from 'node:assert';
import { summarizeProcessTable } from './process-memory-sampler.mjs';

test('sums only browser descendants, grouping renderers and other roles', () => {
  const table = `
  11 1 100 node harness
  12 11 200 node driver
  13 12 300 /path/chrome-headless-shell
  14 13 400 /path/chrome-headless-shell --type=renderer
  15 13 500 /path/chrome-headless-shell --type=renderer
  16 13 600 /path/chrome-headless-shell --type=utility
  17 1 900 /path/chrome-headless-shell --type=renderer
  18 13 900 /some/unrelated/process
`;
  deepStrictEqual(summarizeProcessTable(table, 11), {
    roles: {
      browser: { count: 1, rssBytes: 300 * 1024 },
      renderer: { count: 2, rssBytes: 900 * 1024 },
      utility: { count: 1, rssBytes: 600 * 1024 },
    },
    totalRssBytes: 1800 * 1024,
  });
});

test('handles arbitrary process order and an absent browser', () => {
  deepStrictEqual(
    summarizeProcessTable('3 2 10 /x/chrome-headless-shell --type=gpu-process\n2 1 10 node', 1),
    {
      roles: { gpu: { count: 1, rssBytes: 10240 } },
      totalRssBytes: 10240,
    },
  );
  deepStrictEqual(summarizeProcessTable('2 1 10 node', 1), { roles: {}, totalRssBytes: 0 });
});
