import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/formula-export-browser-benchmark.mjs NEW_REPORT');
const streamPath = resolve('packages/formats/src/stream.ts');
const current = readFileSync(streamPath, 'utf8');
const before = 'book.streamRange(sheet.id, bounds, { ...options, values: !preserve })';
if (current.split(before).length !== 2)
  throw new Error('Cannot reconstruct previous read behavior');
const previous = current.replace(before, 'book.streamRange(sheet.id, bounds, options)');
const variants = {};
for (const [name, contents] of Object.entries({ previous, current })) {
  const bundle = await build({
    entryPoints: [resolve('packages/formats/src/index.ts')],
    bundle: true,
    platform: 'browser',
    format: 'esm',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /formats\/src\/stream\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/formats/src'),
          }));
        },
      },
    ],
  });
  variants[name] = bundle.outputFiles[0].text;
}
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const base = 'http://127.0.0.1:5186/';
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  variantHashes: Object.fromEntries(
    Object.entries(variants).map(([name, text]) => [name, hash(text)]),
  ),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Production preview with hashed assets and separately bundled variants; previous reconstructs formula calculation during read, current requests inputs only. 50,000 simple formulas, three fresh contexts per browser with alternating order. Every emitted cell and report count checked outside timing. Includes Worker read, formatting and output accumulation; excludes XLSX, input latency and Worker/process peak memory.',
  artifacts: [],
  results: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
for (const file of [
  'index.html',
  ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
]) {
  const response = await fetch(new URL(file, base)),
    digest = hash(new Uint8Array(await response.arrayBuffer()));
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
        for (const [name, body] of Object.entries(variants))
          await page.route(`**/formula-export-${name}.js`, (route) =>
            route.fulfill({ contentType: 'text/javascript', body }),
          );
        await page.goto(base + '#budget');
        await page.waitForFunction(() => !!window.opensheet?.selection.get());
        const samples = await page.evaluate(
          async ({ file, round }) => {
            const app = window.opensheet;
            await app.load(file);
            const book = app.getWorkbook(),
              sheet = book.getSheets()[0],
              rows = 5000,
              columns = 10;
            for (let start = 0; start < rows; start += 1000)
              await sheet
                .range({
                  startRow: start,
                  endRow: start + 1000,
                  startColumn: 0,
                  endColumn: columns,
                })
                .setFormulas(Array.from({ length: 1000 }, () => Array(columns).fill('1+1')));
            const out = [];
            for (const variant of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
              const module = await import(`/formula-export-${variant}.js`);
              let exportReport;
              const chunks = [];
              const start = performance.now();
              for await (const chunk of module.streamExport(
                book,
                {
                  sheetId: sheet.id,
                  range: 'A1:J5000',
                  format: 'csv',
                  options: { safe: false, formulas: 'preserve' },
                },
                {
                  onReport: (result) => {
                    exportReport = result;
                  },
                },
              ))
                chunks.push(chunk);
              const elapsedMs = performance.now() - start;
              const lines = chunks.join('').split('\r\n');
              if (lines.length !== rows) throw new Error('Incorrect row count');
              for (const [row, line] of lines.entries()) {
                const cells = line.split(',');
                if (cells.length !== columns || cells.some((value) => value !== '=1+1'))
                  throw new Error(`Incorrect row ${row}`);
              }
              if (
                exportReport?.preservedFormulas !== rows * columns ||
                exportReport.sanitized !== 0 ||
                exportReport.revision !== book.revision
              )
                throw new Error('Incorrect report');
              out.push({
                variant,
                elapsedMs,
                verified: rows * columns,
                preservedFormulas: exportReport.preservedFormulas,
              });
            }
            return out;
          },
          {
            file: createWorkbookFile({ sheets: [{ name: 'Formulas', rows: 5000, columns: 10 }] }),
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
