import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/filter-worker-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: ['packages/core/src/filter.ts'],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  globalName: 'FilterBenchmark',
  write: false,
});
const workerScript =
  bundle.outputFiles[0].text +
  `\nself.onmessage = ({data}) => {
    const query = data.long ? 'X'.repeat(999) + 'Y' : 'AL';
    const criterion = {column: 0, operator: 'contains', value: query};
    const folded = query.toLowerCase();
    const values = data.long ? [query, 'unrelated', 7, null] : ['Alpha', 'beta', 7, null];
    let matches = 0;
    const start = performance.now();
    for (let i = 0; i < 1000000; i++) {
      if (FilterBenchmark.matchesFilter(values[i % values.length], criterion, data.prepared ? folded : undefined)) matches++;
    }
    const ms = performance.now() - start;
    if (matches !== 250000) throw new Error('Incorrect matches: ' + matches);
    postMessage({ms, matches});
  };`;
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  workerSha256: createHash('sha256').update(workerScript).digest('hex'),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless dedicated Worker predicate microbenchmark; one million contains checks per sample, repeating matching text, nonmatching text, number and null. Compared with and without a pre-folded query, 2 or 1000 characters; six alternating rounds with fresh workers. Excludes row reads, calculation, storage, rendering, filtering setup, memory and other devices; does not establish end-to-end performance.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    for (let round = 1; round <= 6; round++) {
      const entry = { round, samples: [] };
      result.rounds.push(entry);
      try {
        entry.samples = await page.evaluate(
          async ({ script, round }) => {
            const samples = [];
            for (const prepared of round % 2 ? [false, true] : [true, false]) {
              for (const long of round % 2 ? [false, true] : [true, false]) {
                const url = URL.createObjectURL(new Blob([script], { type: 'text/javascript' }));
                const worker = new Worker(url);
                try {
                  const response = await new Promise((resolve, reject) => {
                    worker.onmessage = ({ data }) => resolve(data);
                    worker.onerror = reject;
                    worker.postMessage({ prepared, long });
                  });
                  samples.push({ prepared, long, ...response });
                } finally {
                  worker.terminate();
                  URL.revokeObjectURL(url);
                }
              }
            }
            return samples;
          },
          { script: workerScript, round },
        );
        console.log(JSON.stringify({ engine, round, samples: entry.samples }));
      } catch (error) {
        entry.error = String(error);
        throw error;
      } finally {
        save();
      }
    }
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
