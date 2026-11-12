import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/merged-export-browser-benchmark.mjs NEW_REPORT');
const variants = {};
for (const name of ['linear', 'indexed']) {
  const result = await build({
    entryPoints: [resolve('packages/formats/src/index.ts')],
    bundle: true,
    platform: 'browser',
    format: 'esm',
    write: false,
    plugins:
      name === 'linear'
        ? [
            {
              name: 'previous-lookup',
              setup(builder) {
                builder.onLoad({ filter: /generators\/merge-index\.ts$/ }, () => ({
                  loader: 'ts',
                  contents: `export function mergeLookup(merges) { return (r, c) => merges.find(m => r >= m.startRow && r < m.endRow && c >= m.startColumn && c < m.endColumn); }`,
                }));
              },
            },
          ]
        : [],
  });
  variants[name] = result.outputFiles[0].text;
}
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const base = 'http://127.0.0.1:5186/';
const file = createWorkbookFile({ sheets: [{ name: 'Merged', rows: 1000, columns: 64 }] });
file.sheets[0].merges = Array.from({ length: 1000 }, (_, row) => ({
  startRow: row,
  endRow: row + 1,
  startColumn: 0,
  endColumn: 2,
})).reverse();
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  variantHashes: Object.fromEntries(
    Object.entries(variants).map(([name, body]) => [name, hash(body)]),
  ),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Local headless production preview, asset hashes verified. Separately bundled variants reconstruct only previous linear merge lookup. 64,000 blank cells, 1,000 horizontal merges, 3 fresh contexts per browser, alternating variant and format order. Times include Worker reads, rendering and output accumulation; output equality and complete HTML structure checked after timing. No memory, input latency, rich contents, other merge geometries or device claims.',
  artifacts: [],
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const path of [
  'index.html',
  ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
]) {
  const response = await fetch(new URL(path, base)),
    digest = hash(new Uint8Array(await response.arrayBuffer()));
  if (!response.ok || digest !== hash(readFileSync(`apps/playground/dist/${path}`)))
    throw new Error(`Build mismatch: ${path}`);
  report.artifacts.push({ path, sha256: digest });
}
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    for (let round = 1; round <= 3; round++) {
      const context = await browser.newContext();
      const entry = { round, samples: [] };
      result.rounds.push(entry);
      try {
        const page = await context.newPage();
        for (const [name, body] of Object.entries(variants))
          await page.route(`**/merge-bench-${name}.js`, (route) =>
            route.fulfill({ contentType: 'text/javascript', body }),
          );
        await page.goto(base + '#budget');
        await page.waitForFunction(() => !!window.opensheet?.selection.get());
        entry.samples = await page.evaluate(
          async ({ file, round }) => {
            await window.opensheet.load(file);
            const book = window.opensheet.getWorkbook(),
              sheet = book.getSheets()[0];
            const outputs = {},
              samples = [];
            for (const format of round % 2 ? ['html', 'latex'] : ['latex', 'html']) {
              for (const variant of round % 2 ? ['linear', 'indexed'] : ['indexed', 'linear']) {
                const module = await import(`/merge-bench-${variant}.js`);
                const chunks = [],
                  start = performance.now();
                for await (const chunk of module.streamExport(book, {
                  sheetId: sheet.id,
                  range: 'A1:BL1000',
                  format,
                }))
                  chunks.push(chunk);
                const ms = performance.now() - start;
                outputs[`${format}-${variant}`] = chunks.join('');
                samples.push({ format, variant, ms });
              }
              if (outputs[`${format}-linear`] !== outputs[`${format}-indexed`])
                throw new Error(`${format} variant output mismatch`);
            }
            const doc = new DOMParser().parseFromString(outputs['html-indexed'], 'text/html'),
              rows = [...doc.querySelectorAll('tr')];
            if (rows.length !== 1000) throw new Error('Incorrect HTML row count');
            for (const [i, row] of rows.entries()) {
              const cells = [...row.children];
              if (
                cells.length !== 63 ||
                cells[0].getAttribute('colspan') !== '2' ||
                cells.some((cell) => cell.textContent !== '')
              )
                throw new Error(`Invalid HTML row ${i}`);
            }
            if ((outputs['latex-indexed'].match(/\\multicolumn\{2\}/g) ?? []).length !== 1000)
              throw new Error('Incorrect LaTeX span count');
            return samples.map((sample) => ({ ...sample, verifiedRows: rows.length }));
          },
          { file, round },
        );
        console.log(JSON.stringify({ engine, round, samples: entry.samples }));
      } catch (error) {
        entry.error = String(error);
        throw error;
      } finally {
        save();
        await context.close();
      }
    }
  } finally {
    save();
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
