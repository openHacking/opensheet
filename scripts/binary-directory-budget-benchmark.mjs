import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { createHash } from 'node:crypto';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output)
  throw new Error('Usage: node scripts/binary-directory-budget-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: ['packages/core/src/binary-file.ts'],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  globalName: 'BinaryBudgetBenchmark',
  write: false,
});
const content = bundle.outputFiles[0].text;
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  bundleSha256: createHash('sha256').update(content).digest('hex'),
  scope:
    'Paired synthetic in-page directory construction and serialization. Control omits the new early budget check; it is not a previous commit. All 100,000 generated entries and serialized output are verified after timing. Excludes page encoding, Worker transfer, storage, process memory and reference devices.',
  browsers: {},
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [name, engine] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await engine.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.addScriptTag({ content });
    const samples = await page.evaluate(() => {
      const expected = Array.from({ length: 100000 }, (_, i) => ({
        sheet: i % 4,
        row: i,
        column: i % 3125,
        offset: 1024 + i * 1024,
        length: 1024,
      }));
      const expectedJSON = JSON.stringify(expected);
      const results = [];
      for (let round = 0; round < 6; round++)
        for (const variant of round % 2 ? ['checked', 'control'] : ['control', 'checked']) {
          const start = performance.now();
          const entries = [];
          const budget =
            variant === 'checked'
              ? new globalThis.BinaryBudgetBenchmark.BinaryDirectoryBudget()
              : undefined;
          for (const entry of expected) {
            budget?.add(entry);
            entries.push(entry);
          }
          const encoded = JSON.stringify(entries);
          const ms = performance.now() - start;
          if (encoded !== expectedJSON || entries.length !== expected.length)
            throw new Error('Directory differs');
          results.push({
            round,
            variant,
            ms,
            entries: entries.length,
            encodedBytes: encoded.length,
          });
        }
      return results;
    });
    report.browsers[name] = { version: browser.version(), samples };
    save();
    console.log(name, 'completed', samples.length, 'samples');
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
