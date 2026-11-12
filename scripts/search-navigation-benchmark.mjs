import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/search-navigation-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const report = {
  engine,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  rounds: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5186/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(async () => {
        const app = window.opensheet;
        let book = app.getWorkbook();
        await book.generate(1_000_000);
        const query = 'opensheet-benchmark-needle';
        for (const address of ['A1', 'A5000', 'A10000'])
          await book
            .getSheets()[0]
            .range(address)
            .setValues([[query]]);
        await app.open(book.id);
        book = app.getWorkbook();
        const sheet = book.getSheets()[0];
        const samples = [];
        const search = async (after, direction, row, count) => {
          const start = performance.now();
          const result = await book.search(sheet.id, query, after, { direction });
          const elapsedMs = performance.now() - start;
          if (result.count !== count || result.match?.row !== row || result.match?.column !== 0)
            throw new Error(`Unexpected search result ${JSON.stringify(result)}`);
          return elapsedMs;
        };
        const coldMs = await search(-1, 'next', 0, 3);
        const positions = [0, 4999, 9999];
        let current = 0;
        for (let step = 0; step < 30; step++) {
          const backward = step % 4 < 2;
          const next = (current + (backward ? 2 : 1)) % 3;
          samples.push(
            await search(
              positions[current] * 100,
              backward ? 'previous' : 'next',
              positions[next],
              3,
            ),
          );
          current = next;
        }
        await sheet.range('A5000').setValues([['removed']]);
        const afterEditMs = await search(0, 'next', 9999, 2);
        return { cells: 1_000_000, coldMs, navigationMs: samples, afterEditMs };
      });
      report.rounds.push({ round, ...result });
      console.log(
        JSON.stringify({
          engine,
          round,
          coldMs: result.coldMs,
          afterEditMs: result.afterEditMs,
          warmMaxMs: Math.max(...result.navigationMs),
        }),
      );
    } catch (error) {
      report.rounds.push({ round, error: String(error) });
      throw error;
    } finally {
      save();
      await context.close();
    }
  }
} finally {
  await browser.close();
}
