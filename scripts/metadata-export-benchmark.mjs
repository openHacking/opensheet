import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/metadata-export-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const file = createWorkbookFile({ sheets: [{ name: 'Metadata', rows: 100_000, columns: 1 }] });
file.sheets[0].rows.meta = Object.fromEntries(
  Array.from({ length: 100_000 }, (_, row) => [
    String(row),
    { size: 20 + (row % 10), hidden: row % 2 === 0 },
  ]),
);
file.extensions.fixture = { label: 'original' };
const report = {
  engine,
  startedAt: new Date().toISOString(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  limitations:
    'Headless local single-host samples. Metadata-only fixture; no cell pages, matched pre-change baseline, process memory or cross-device claims. Standalone clone timing is not isolated toJSON overhead.',
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
      const result = await page.evaluate(async (input) => {
        const app = window.opensheet;
        await app.load(input);
        const book = app.getWorkbook();
        const start = performance.now();
        const exported = await book.toJSON();
        const exportMs = performance.now() - start;
        const cloneStart = performance.now();
        const copy = structuredClone(exported);
        const cloneMs = performance.now() - cloneStart;
        const rows = copy.sheets[0].rows.meta;
        if (Object.keys(rows).length !== 100_000) throw new Error('Incorrect row metadata count');
        for (let row = 0; row < 100_000; row++)
          if (rows[row]?.size !== 20 + (row % 10) || rows[row]?.hidden !== (row % 2 === 0))
            throw new Error(`Incorrect row metadata: ${row}`);
        exported.extensions.fixture.label = 'modified';
        exported.sheets[0].rows.meta['0'].size = 999;
        const fresh = await book.toJSON();
        if (
          fresh.extensions.fixture.label !== 'original' ||
          fresh.sheets[0].rows.meta['0'].size !== 20
        )
          throw new Error('Metadata aliases live workbook');
        return { rows: 100_000, exportMs, cloneMs };
      }, file);
      report.rounds.push({ round, ...result });
      console.log(JSON.stringify({ engine, round, ...result }));
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
