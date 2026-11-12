import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/stream-report-browser-benchmark.mjs NEW_REPORT');
const bundled = await build({
  entryPoints: [resolve('packages/formats/src/index.ts')],
  bundle: true,
  platform: 'browser',
  format: 'esm',
  write: false,
});
const harness = bundled.outputFiles[0].text;
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const base = 'http://127.0.0.1:5186/';
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  harnessSha256: hash(harness),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Production preview with matching served assets and a separately bundled current formats module; 100,000 synthetic unsafe text cells; alternating report/no-report streams in three fresh contexts per browser. Full output checked outside timed interval. Timing includes Worker reads, formatting, text accumulation and report callback. No XLSX, input latency or Worker/process peak memory measurement.',
  artifacts: [],
  results: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
for (const file of [
  'index.html',
  ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
]) {
  const response = await fetch(new URL(file, base));
  const digest = hash(new Uint8Array(await response.arrayBuffer()));
  if (!response.ok || digest !== hash(readFileSync(`apps/playground/dist/${file}`)))
    throw new Error(`Build mismatch: ${file}`);
  report.artifacts.push({ file, sha256: digest });
}
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  try {
    const result = { engine, version: browser.version(), rounds: [] };
    report.results.push(result);
    for (let round = 0; round < 3; round++) {
      const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
      try {
        const page = await context.newPage();
        await page.route('**/stream-report-benchmark.js', (route) =>
          route.fulfill({ contentType: 'text/javascript', body: harness }),
        );
        await page.goto(base + '#budget');
        await page.waitForFunction(() => !!window.opensheet?.selection.get());
        const samples = await page.evaluate(
          async ({ file, round }) => {
            const { streamExport } = await import('/stream-report-benchmark.js');
            const app = window.opensheet;
            await app.load(file);
            const book = app.getWorkbook(),
              sheet = book.getSheets()[0];
            const rows = 10_000,
              columns = 10;
            for (let start = 0; start < rows; start += 1000) {
              const values = Array.from({ length: 1000 }, (_, r) =>
                Array.from({ length: columns }, (_, c) => `=${start + r}-${c}`),
              );
              await sheet
                .range({
                  startRow: start,
                  endRow: start + 1000,
                  startColumn: 0,
                  endColumn: columns,
                })
                .setValues(values);
            }
            const config = { sheetId: sheet.id, range: 'A1:J10000', format: 'csv' };
            const out = [];
            for (const variant of round % 2 ? ['report', 'no-report'] : ['no-report', 'report']) {
              let value;
              const chunks = [];
              const start = performance.now();
              for await (const chunk of streamExport(
                book,
                config,
                variant === 'report'
                  ? {
                      onReport: (result) => {
                        value = result;
                      },
                    }
                  : {},
              ))
                chunks.push(chunk);
              const elapsedMs = performance.now() - start;
              const lines = chunks.join('').split('\r\n');
              if (lines.length !== rows) throw new Error('Unexpected row count');
              for (let row = 0; row < rows; row++) {
                const cells = lines[row].split(',');
                if (cells.length !== columns) throw new Error(`Unexpected width ${row}`);
                for (let col = 0; col < columns; col++)
                  if (cells[col] !== `'=${row}-${col}`)
                    throw new Error(`Incorrect cell ${row},${col}`);
              }
              const expectedSample = Array.from(
                { length: 32 },
                (_, i) =>
                  `${String.fromCharCode(65 + (i % columns))}${Math.floor(i / columns) + 1}`,
              );
              if (
                variant === 'report' &&
                (value?.format !== 'csv' ||
                  value.sanitized !== rows * columns ||
                  !value.truncated ||
                  JSON.stringify(value.sample) !== JSON.stringify(expectedSample) ||
                  value.revision !== book.revision)
              )
                throw new Error('Incorrect report');
              out.push({ variant, elapsedMs, verified: rows * columns, report: value ?? null });
            }
            return out;
          },
          {
            file: createWorkbookFile({ sheets: [{ name: 'Report', rows: 10_000, columns: 10 }] }),
            round,
          },
        );
        result.rounds.push({ round: round + 1, samples });
        save();
        console.log(JSON.stringify({ engine, round: round + 1, samples }));
      } finally {
        await context.close();
      }
    }
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
