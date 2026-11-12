import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync, readFileSync, readdirSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, totalmem, platform, arch, release } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new output path');
const styleCount = Number(process.argv.find((arg) => arg.startsWith('--styles='))?.slice(9) ?? 0);
if (!Number.isSafeInteger(styleCount) || styleCount < 0 || styleCount > 10000)
  throw new Error('Use 0–10000 styles');
function fixture() {
  const file = createWorkbookFile({
    sheets: [
      { name: 'Paste', rows: 100, columns: 50 },
      ...(styleCount ? [{ name: 'Styles', rows: styleCount, columns: 1 }] : []),
    ],
  });
  if (!styleCount) return file;
  file.styles = Array.from({ length: styleCount }, (_, i) => ({
    color: `#${i.toString(16).padStart(6, '0')}`,
  }));
  file.sheets[1].blocks = Array.from({ length: Math.ceil(styleCount / 64) }, (_, block) => {
    const count = Math.min(64, styleCount - block * 64);
    return {
      row: block,
      column: 0,
      data: {
        blanks: Array.from({ length: count }, (_, row) => row * 32),
        styles: Array.from({ length: count }, (_, row) => [row * 32, 1, block * 64 + row]),
      },
    };
  });
  return file;
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  host: {
    node: process.version,
    cpu: cpus()[0].model,
    memoryBytes: totalmem(),
    platform: platform(),
    arch: arch(),
    release: release(),
    powerMode: 'unrecorded',
  },
  scope:
    'Production preview; 5,000 numeric input-only cells; durable command latency excludes clipboard parsing and paint. Commands alternate within each fresh workbook; no merged cells. Full values and undo checked. Three sequential contexts per browser, four paired measurements per context. Not an interaction latency or cross-device claim.',
  styleCount,
  results: [],
};
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  const base = 'http://127.0.0.1:5186/';
  const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
  report.artifacts = [];
  for (const file of [
    'index.html',
    ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
  ]) {
    const response = await fetch(new URL(file, base));
    const sha256 = hash(new Uint8Array(await response.arrayBuffer()));
    if (!response.ok || sha256 !== hash(readFileSync(`apps/playground/dist/${file}`)))
      throw new Error(`Build mismatch: ${file}`);
    report.artifacts.push({ file, sha256 });
  }
  for (const [engine, launcher] of Object.entries({ chromium, webkit, firefox })) {
    const browser = await launcher.launch();
    try {
      const result = { engine, version: browser.version(), rounds: [] };
      report.results.push(result);
      save();
      for (let round = 0; round < 3; round++) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
        try {
          const page = await context.newPage();
          await page.goto(base + '#budget');
          await page.waitForFunction(() => !!window.opensheet?.selection.get());
          const samples = await page.evaluate(
            async ({ file, round, styleCount }) => {
              const app = window.opensheet;
              await app.load(file);
              const book = app.getWorkbook(),
                sheet = book.getSheets()[0];
              const range = { startRow: 0, endRow: 100, startColumn: 0, endColumn: 50 };
              const cells = Array.from({ length: 5000 }, (_, i) => ({
                row: Math.floor(i / 50),
                column: i % 50,
                input: { type: 'number', value: i + 1 },
              }));
              if (Object.keys(book.metadata.styles).length !== styleCount)
                throw new Error('Metadata style count mismatch');
              const out = [];
              for (let pair = 0; pair < 4; pair++) {
                const order =
                  (pair + round) % 2
                    ? ['core.cells.paste', 'core.cells.set']
                    : ['core.cells.set', 'core.cells.paste'];
                for (const type of order) {
                  const start = performance.now();
                  await book.execute({ type, payload: { sheetId: sheet.id, cells } });
                  const durationMs = performance.now() - start;
                  const values = await sheet.range(range).getValues();
                  if (
                    values.length !== 100 ||
                    values.some(
                      (row, r) => row.length !== 50 || row.some((v, c) => v !== r * 50 + c + 1),
                    )
                  )
                    throw new Error('Written values mismatch');
                  if (!(await book.undo())) throw new Error('Undo failed');
                  const undone = await sheet.range(range).getValues();
                  if (undone.some((row) => row.some((v) => v !== null)))
                    throw new Error('Undo values mismatch');
                  out.push({ pair, type, durationMs });
                }
              }
              const id = book.id;
              await app.open(id);
              const reopened = await app.getWorkbook().getSheets()[0].range(range).getValues();
              if (reopened.some((row) => row.some((v) => v !== null)))
                throw new Error('Reopen mismatch');
              if (Object.keys(app.getWorkbook().metadata.styles).length !== styleCount)
                throw new Error('Reopened style count mismatch');
              return out;
            },
            {
              file: fixture(),
              round,
              styleCount,
            },
          );
          result.rounds.push({ round, samples });
          save();
          process.stdout.write(JSON.stringify({ engine, round, samples }) + '\n');
        } finally {
          await context.close();
        }
      }
    } finally {
      await browser.close();
    }
  }
  if (JSON.stringify(localBuildFingerprint()) !== JSON.stringify(report.build))
    throw new Error('Build changed during measurement');
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
}
