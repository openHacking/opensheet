import { chromium } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { existsSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { arch, platform, cpus, totalmem, release } from 'node:os';
import { execFileSync } from 'node:child_process';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const option = (name, fallback) =>
  process.argv.find((arg) => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback;
const rows = Number(option('rows', '10000'));
const rounds = Number(option('rounds', '3'));
const formulaOnly = option('formula-only', 'false') === 'true';
if (
  !Number.isSafeInteger(rows) ||
  rows < 100 ||
  rows > 50000 ||
  !Number.isSafeInteger(rounds) ||
  rounds < 1 ||
  rounds > 10
)
  throw new Error('Use 100–50000 rows and 1–10 rounds');
const output = resolve(option('output', 'benchmarks/data-operations.json'));
if (existsSync(output)) throw new Error('Refusing to overwrite benchmark evidence');
const report = {
  startedAt: new Date().toISOString(),
  rows,
  columns: 4,
  rounds,
  formulaOnly,
  url: option('url', 'http://127.0.0.1:5186/'),
  sourceCommit: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
  sourceDirty: Boolean(execFileSync('git', ['status', '--porcelain'], { encoding: 'utf8' }).trim()),
  sourceFingerprint: sourceFingerprint(),
  localBuildFingerprint: localBuildFingerprint(),
  platform: platform(),
  arch: arch(),
  osRelease: release(),
  node: process.version,
  hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length, memoryBytes: totalmem() },
  powerMode: option('power-mode', 'unrecorded'),
  headless: true,
  limitations:
    'Total operation timings, not internal phase profiling. Main-document frame/heap samples exclude Worker retained memory. No performance pass threshold is asserted.',
  results: [],
};
mkdirSync(dirname(output), { recursive: true });
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
save();
const browser = await chromium.launch();
report.browserVersion = browser.version();
try {
  for (let round = 1; round <= rounds; round++) {
    const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    try {
      const page = await context.newPage();
      await page.goto(report.url);
      await page.waitForFunction(
        () => !!window.opensheet && !document.querySelector('#scene-select').disabled,
      );
      const result = await page.evaluate(
        async ({ rows, file, formulaOnly }) => {
          const app = window.opensheet;
          await app.load(file);
          let book = app.getWorkbook(),
            sheet = book.getSheets()[0];
          const sortPhases = [];
          book.onProgress((progress) => {
            if (progress.phase.startsWith('sort-') && progress.elapsedMs !== undefined)
              sortPhases.push(progress);
          });
          const timings = {},
            frames = [];
          let previous = performance.now(),
            active = true,
            frameId;
          const tick = (now) => {
            frames.push(now - previous);
            previous = now;
            if (active) frameId = requestAnimationFrame(tick);
          };
          frameId = requestAnimationFrame(tick);
          const timed = async (name, action) => {
            const start = performance.now();
            const result = await action();
            timings[name] = performance.now() - start;
            return result;
          };
          const assert = (condition, message) => {
            if (!condition) throw new Error(message);
          };
          const source = Array.from({ length: rows }, (_, i) => [i % 11, (i * 17) % 97, i]);
          const expected = source.map((row) => [...row, row[1] * 2]);
          const filter = {
            mode: 'all',
            headerRows: 0,
            criteria: [
              ...(!formulaOnly ? [{ column: 0, operator: 'eq', value: 3 }] : []),
              { column: 3, operator: 'gte', value: 100 },
            ],
          };
          const hidden = (data) =>
            data.flatMap((row, i) => ((formulaOnly || row[0] === 3) && row[3] >= 100 ? [] : [i]));
          try {
            await timed('populateMs', async () => {
              await sheet.range(`A1:C${rows}`).setValues(source);
              await sheet.range(`D1:D${rows}`).setFormulas(source.map((_, i) => [`B${i + 1}*2`]));
            });
            await timed('filterMs', () => sheet.setFilter(filter));
            assert(
              JSON.stringify(book.filteredRows(sheet.id)) === JSON.stringify(hidden(expected)),
              'Initial filter mismatch',
            );
            await timed('clearFilterMs', () => sheet.setFilter(null));
            const sorted = [...expected].sort((a, b) => a[0] - b[0] || b[3] - a[3] || a[2] - b[2]);
            await timed('formulaMultiSortMs', () =>
              sheet.sort(`A1:D${rows}`, [
                { column: 0, direction: 'asc' },
                { column: 3, direction: 'desc' },
              ]),
            );
            const verify = async (data) => {
              for (let start = 0; start < rows; start += 4096) {
                const end = Math.min(start + 4096, rows);
                const actual = await sheet
                  .range({ startRow: start, endRow: end, startColumn: 0, endColumn: 4 })
                  .getValues();
                assert(
                  JSON.stringify(actual) === JSON.stringify(data.slice(start, end)),
                  `Full-cell verification mismatch at row ${start}`,
                );
              }
            };
            await timed('sortedVerificationMs', () => verify(sorted));
            await timed('sortedFilterMs', () => sheet.setFilter(filter));
            assert(
              JSON.stringify(book.filteredRows(sheet.id)) === JSON.stringify(hidden(sorted)),
              'Sorted filter mismatch',
            );
            await timed('undoMs', async () => {
              await book.undo();
              await book.undo();
            });
            await verify(expected);
            assert(book.filteredRows(sheet.id).length === 0, 'Undo filter mismatch');
            await timed('redoMs', async () => {
              await book.redo();
              await book.redo();
            });
            await verify(sorted);
            const id = book.id;
            await timed('reopenMs', () => app.open(id));
            book = app.getWorkbook();
            sheet = book.getSheets()[0];
            await timed('reopenedVerificationMs', () => verify(sorted));
            assert(
              JSON.stringify(book.filteredRows(sheet.id)) === JSON.stringify(hidden(sorted)),
              'Reopened filter mismatch',
            );
            const stats = await book.storageStats();
            assert(stats.bytes <= stats.budgetBytes, 'Storage budget exceeded');
            const orderedFrames = [...frames].sort((a, b) => a - b);
            return {
              timings,
              sortPhases,
              verifiedCells: rows * 4,
              verification: 'passed',
              storage: stats,
              frameSamples: frames,
              frameP95Ms: orderedFrames[Math.floor(orderedFrames.length * 0.95)] ?? null,
              sampledMainHeapBytes: performance.memory?.usedJSHeapSize ?? null,
            };
          } finally {
            active = false;
            cancelAnimationFrame(frameId);
          }
        },
        {
          rows,
          formulaOnly,
          file: createWorkbookFile({ sheets: [{ name: 'Data operations', rows, columns: 4 }] }),
        },
      );
      report.results.push({ round, ...result });
      save();
      process.stdout.write(
        JSON.stringify({ round, timings: result.timings, verification: result.verification }) +
          '\n',
      );
    } finally {
      await context.close();
    }
  }
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
} finally {
  await browser.close();
}
