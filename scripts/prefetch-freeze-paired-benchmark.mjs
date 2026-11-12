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
  scriptSha256: hash(readFileSync('scripts/prefetch-freeze-paired-benchmark.mjs')),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Paired development-server standalone workbook, real Worker/IndexedDB, one fully populated 64x32 numeric page and four synchronous listeners. Baseline disables only deep freezing of the returned page. Three alternating pairs of fresh contexts per browser, 15 samples each after three warmups; viewport cache cleared and every value and callback count checked outside timing. Excludes UI paint, peak memory, real devices and other workloads. No claim of complete old-commit equivalence; one host, power mode unrecorded.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const find = 'this.cache.set(a, freeze(data, true));';
const replacement = 'this.cache.set(a, data);';
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
              await sheet
                .range('A1:AF64')
                .setValues(
                  Array.from({ length: 64 }, (_, r) =>
                    Array.from({ length: 32 }, (_, c) => r * 32 + c),
                  ),
                );
              let calls = 0;
              const stops = Array.from({ length: 4 }, () =>
                book.onData(() => {
                  calls++;
                }),
              );
              const samples = [];
              for (let i = 0; i < 18; i++) {
                book.cache.clear();
                const before = calls,
                  start = performance.now();
                await book.prefetch(sheet.id, range);
                const ms = performance.now() - start;
                if (calls - before !== 4) throw new Error('Unexpected data or callback count');
                for (let row = 0; row < 64; row++)
                  for (let col = 0; col < 32; col++)
                    if (book.peekValue(sheet.id, row, col) !== row * 32 + col)
                      throw new Error(`Wrong value ${row}:${col}`);
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
