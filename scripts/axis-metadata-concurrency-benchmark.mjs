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
    .update(readFileSync('scripts/axis-metadata-concurrency-benchmark.mjs'))
    .digest('hex'),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Development-server real Worker/IndexedDB workbook with 100,000 row size records, one context per browser. Three alternating pairs toggle one row then perform three no-op column visibility commands, queued concurrently or awaited sequentially. Times include Worker, response transport and client adoption; setup, validation and JSON character counting are excluded. JSON length is not wire bytes. Not a comparison with the old protocol; no peak memory, rendering, reference device or power-mode measurement.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  try {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5173/');
      const entry = { engine, version: browser.version() };
      report.results.push(entry);
      entry.samples = await page.evaluate(async (modulePath) => {
        const { createWorkbook } = await import(/* @vite-ignore */ modulePath);
        const book = await createWorkbook({
          sheets: [{ name: 'Large', rows: 100_000, columns: 8 }],
        });
        try {
          const sheet = book.getSheets()[0];
          await sheet.setRowsHeight(0, 100_000, 42);
          const worker = book.worker;
          const handler = worker.onmessage;
          let files = [];
          worker.onmessage = (event) => {
            const file = event.data.result?.view?.file;
            if (file) files.push(file);
            handler.call(worker, event);
          };
          const samples = [];
          let hidden = false;
          for (let pair = 0; pair < 3; pair++) {
            for (const variant of pair % 2 ? ['sequential', 'queued'] : ['queued', 'sequential']) {
              hidden = !hidden;
              files = [];
              const revision = book.revision;
              const started = performance.now();
              if (variant === 'queued') {
                await Promise.all([
                  sheet.setRowsHidden(0, 1, hidden),
                  sheet.setColumnsHidden(0, 1, false),
                  sheet.setColumnsHidden(1, 1, false),
                  sheet.setColumnsHidden(2, 1, false),
                ]);
              } else {
                await sheet.setRowsHidden(0, 1, hidden);
                for (let column = 0; column < 3; column++)
                  await sheet.setColumnsHidden(column, 1, false);
              }
              const elapsedMs = performance.now() - started;
              const metadataCharacters = files.map((file) => JSON.stringify(file).length);
              const data = book.sheetData(sheet.id);
              if (
                book.revision !== revision + 1 ||
                !!data.rows[data.rowOrder[0]]?.hidden !== hidden
              )
                throw new Error('Incorrect revision or visibility');
              samples.push({ pair: pair + 1, variant, elapsedMs, metadataCharacters });
            }
          }
          return samples;
        } finally {
          await book.deleteStorage();
        }
      }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
      save();
    } finally {
      await context.close();
    }
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
console.log(JSON.stringify(report.results));
