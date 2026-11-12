import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Pass a new report path; start the development server on port 5173');
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  scope:
    'Headless development server, 256 rows; alternating searches with and without AbortSignal, each checking a single match in the final row. Compares current paths, not an old build; excludes memory and other devices.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, browser: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    for (let round = 0; round < 3; round++) {
      const context = await browser.newContext();
      try {
        const page = await context.newPage();
        await page.goto('http://127.0.0.1:5173/');
        await page.waitForFunction(() => !!window.opensheet?.selection.get());
        const sample = await page.evaluate(async (round) => {
          const book = window.opensheet.getWorkbook();
          const longSheet = await book.addSheet('Long text', { rows: 256, columns: 1 });
          const shortSheet = await book.addSheet('Short text', { rows: 256, columns: 1 });
          const markers = Array.from({ length: 12 }, (_, i) => `|needle${i}|`).join('');
          await longSheet
            .range('A1:A256')
            .setValues(
              Array.from({ length: 256 }, (_, i) => [
                i === 255 ? 'x'.repeat(32500) + markers : 'x'.repeat(32760),
              ]),
            );
          await shortSheet
            .range('A1:A256')
            .setValues(Array.from({ length: 256 }, (_, i) => [i === 255 ? markers : 'short']));
          const measurements = [];
          for (const [name, sheet] of [
            ['long', longSheet],
            ['short', shortSheet],
          ]) {
            for (let pair = 0; pair < 6; pair++) {
              const modes = (pair + round) % 2 ? [true, false] : [false, true];
              for (const cooperative of modes) {
                const signal = cooperative ? new AbortController().signal : undefined;
                const start = performance.now();
                const found = await book.search(
                  sheet.id,
                  `|needle${pair * 2 + (cooperative ? 1 : 0)}|`,
                  -1,
                  { signal },
                );
                const ms = performance.now() - start;
                if (found.count !== 1 || found.match?.row !== 255 || found.match?.column !== 0)
                  throw new Error('Unexpected search result');
                measurements.push({ name, cooperative, pair, ms });
              }
            }
          }
          return measurements;
        }, round);
        result.rounds.push({ round: round + 1, samples: sample });
      } catch (error) {
        result.rounds.push({ round: round + 1, error: String(error) });
        throw error;
      } finally {
        await context.close();
        save();
      }
    }
  } finally {
    await browser.close();
  }
}
