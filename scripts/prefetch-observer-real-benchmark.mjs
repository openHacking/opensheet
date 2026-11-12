import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  scriptSha256: createHash('sha256')
    .update(readFileSync('scripts/prefetch-observer-real-benchmark.mjs'))
    .digest('hex'),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Development-server standalone workbook; real Worker/IndexedDB prefetch of one 64x32 page after clearing the main viewport cache. Three fresh browser contexts per engine, 20 samples per 0/1/4 synchronous data listeners in alternating order; every result and callback counted. Includes Worker, transport and cache refill, excludes UI paint, peak memory, failure and async observers; no old implementation baseline, one machine and unrecorded power mode.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const entry = { engine, version: browser.version(), rounds: [] };
  report.results.push(entry);
  try {
    for (let round = 0; round < 3; round++) {
      const context = await browser.newContext();
      try {
        const page = await context.newPage();
        await page.goto('http://127.0.0.1:5173/');
        const modulePath = `/@fs${process.cwd()}/packages/core/src/index.ts`;
        const samples = await page.evaluate(
          async ({ modulePath, round }) => {
            const { createWorkbook } = await import(/* @vite-ignore */ modulePath);
            const book = await createWorkbook({
              sheets: [{ name: 'Prefetch', rows: 64, columns: 32 }],
            });
            const sheet = book.getSheets()[0];
            const range = sheet.range('A1').bounds;
            await sheet.range('A1').setValues([[12345]]);
            const results = [];
            try {
              for (const count of round % 2 ? [4, 1, 0] : [0, 1, 4]) {
                let calls = 0;
                const stops = Array.from({ length: count }, () =>
                  book.onData(() => {
                    calls++;
                  }),
                );
                const samples = [];
                for (let i = 0; i < 23; i++) {
                  book.cache.clear();
                  const before = calls,
                    start = performance.now();
                  await book.prefetch(sheet.id, range);
                  const ms = performance.now() - start;
                  if (book.peekValue(sheet.id, 0, 0) !== 12345 || calls - before !== count)
                    throw new Error(
                      `Unexpected prefetch result or callbacks: ${calls - before}/${count}`,
                    );
                  if (i >= 3) samples.push(ms);
                }
                stops.forEach((stop) => stop());
                results.push({ count, callbacks: calls, samplesMs: samples });
              }
            } finally {
              await book.deleteStorage();
            }
            return results;
          },
          { modulePath, round },
        );
        entry.rounds.push({ round: round + 1, samples });
        save();
      } finally {
        await context.close();
      }
    }
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
console.log(
  JSON.stringify(report.results.map(({ engine, rounds }) => ({ engine, rounds: rounds.length }))),
);
