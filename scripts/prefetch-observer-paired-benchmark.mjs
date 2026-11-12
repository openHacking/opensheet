import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const hash = (value) => createHash('sha256').update(value).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  scriptSha256: hash(readFileSync('scripts/prefetch-observer-paired-benchmark.mjs')),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Paired development-server standalone workbook, real Worker/IndexedDB, one 64x32 page and four nonthrowing synchronous listeners. The baseline changes only the served prefetch notification loop back to the former direct call; all other code remains current. Three alternating pairs of fresh contexts per browser, 40 samples each after three warmups; viewport cache cleared, result and callback count checked on every sample. Excludes failures, asynchronous listeners, UI paint, peak memory, real devices and other workloads. No claim of complete old-commit equivalence; one host, power mode unrecorded.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const find = 'for (const fn of this.dataListeners) notifyObserver(fn);';
const replacement = 'for (const fn of this.dataListeners) fn();';
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const entry = { engine, version: browser.version(), pairs: [] };
  report.results.push(entry);
  try {
    for (let round = 0; round < 3; round++) {
      const pair = { round: round + 1, variants: [] };
      entry.pairs.push(pair);
      for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
        const context = await browser.newContext();
        try {
          const page = await context.newPage();
          const hashes = [];
          await page.route('**/packages/core/src/workbook.ts*', async (route) => {
            const response = await route.fetch();
            const original = await response.text();
            if (original.split(find).length !== 2)
              throw new Error('Expected exactly one prefetch notification loop');
            const body = variant === 'baseline' ? original.replace(find, replacement) : original;
            hashes.push({ original: hash(original), served: hash(body) });
            await route.fulfill({ response, body });
          });
          await page.goto('http://127.0.0.1:5173/');
          const modulePath = `/@fs${process.cwd()}/packages/core/src/index.ts`;
          const samples = await page.evaluate(async (modulePath) => {
            const { createWorkbook } = await import(/* @vite-ignore */ modulePath);
            const book = await createWorkbook({
              sheets: [{ name: 'Prefetch', rows: 64, columns: 32 }],
            });
            try {
              const sheet = book.getSheets()[0],
                range = sheet.range('A1').bounds;
              await sheet.range('A1').setValues([[12345]]);
              let calls = 0;
              const stops = Array.from({ length: 4 }, () =>
                book.onData(() => {
                  calls++;
                }),
              );
              const samples = [];
              for (let i = 0; i < 43; i++) {
                book.cache.clear();
                const before = calls,
                  start = performance.now();
                await book.prefetch(sheet.id, range);
                const ms = performance.now() - start;
                if (calls - before !== 4 || book.peekValue(sheet.id, 0, 0) !== 12345)
                  throw new Error('Unexpected data or callback count');
                if (i >= 3) samples.push(ms);
              }
              stops.forEach((stop) => stop());
              return { calls, samples };
            } finally {
              await book.deleteStorage();
            }
          }, modulePath);
          if (!hashes.length) throw new Error('Workbook module route not observed');
          pair.variants.push({ variant, hashes, ...samples });
          save();
        } finally {
          await context.close();
        }
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
  JSON.stringify(report.results.map(({ engine, pairs }) => ({ engine, pairs: pairs.length }))),
);
