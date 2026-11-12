import { chromium, firefox, webkit } from '@playwright/test';
import { transform } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const source = readFileSync('packages/core/src/workbook.ts', 'utf8');
const start = source.indexOf('function notifyObserver<');
const end = source.indexOf('\nexport interface OperationTiming', start);
if (start < 0 || end < 0) throw new Error('Observer source not found');
const helperSource = source.slice(start, end);
const helper = (await transform(helperSource, { loader: 'ts', target: 'es2022' })).code;
const hash = (value) => createHash('sha256').update(value).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  helperSourceSha256: hash(helperSource),
  benchmarkSha256: hash(readFileSync('scripts/prefetch-observer-benchmark.mjs')),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless main-page callback-dispatch microbenchmark. Compares the old direct prefetch data-listener loop with the current extracted notification helper, using 100,000 synthetic notifications and zero, one or four synchronous nonthrowing listeners. Seven alternating samples per profile after warmup; checks every callback count. Not an end-to-end workbook prefetch, Worker/IndexedDB, paint, memory, asynchronous or failure benchmark. One host, no recorded power mode.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const entry = { engine, version: browser.version(), samples: [] };
  report.results.push(entry);
  try {
    const page = await browser.newPage();
    entry.samples = await page.evaluate(
      ({ helper }) => {
        const notifyObserver = new Function(`${helper}\nreturn notifyObserver`)();
        const rounds = 100_000;
        const profiles = [0, 1, 4];
        const run = (count, current) => {
          let calls = 0;
          const listeners = Array.from({ length: count }, () => () => {
            calls++;
          });
          const start = performance.now();
          for (let i = 0; i < rounds; i++) {
            if (current) for (const listener of listeners) notifyObserver(listener);
            else for (const listener of listeners) listener();
          }
          const ms = performance.now() - start;
          if (calls !== count * rounds) throw new Error(`Unexpected callback count ${calls}`);
          return { count, current, calls, ms };
        };
        for (const count of profiles)
          for (let i = 0; i < 3; i++) {
            run(count, false);
            run(count, true);
          }
        const samples = [];
        for (let round = 0; round < 7; round++) {
          for (const count of round % 2 ? [...profiles].reverse() : profiles) {
            for (const current of round % 2 ? [true, false] : [false, true])
              samples.push({ round: round + 1, ...run(count, current) });
          }
        }
        return samples;
      },
      { helper },
    );
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
console.log(
  JSON.stringify(
    report.results.map(({ engine, samples }) => ({ engine, samples: samples.length })),
  ),
);
