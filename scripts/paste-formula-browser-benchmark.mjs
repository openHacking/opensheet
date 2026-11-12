import { chromium, firefox, webkit } from '@playwright/test';
import { readFileSync, readdirSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, release } from 'node:os';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium', rowArgument = '200'] = process.argv.slice(2);
const rowCount = Number(rowArgument);
if (
  !output ||
  !['chromium', 'firefox', 'webkit'].includes(engine) ||
  !Number.isInteger(rowCount) ||
  rowCount < 100 ||
  rowCount > 2000 ||
  rowCount % 100
)
  throw new Error(
    'Usage: node scripts/paste-formula-browser-benchmark.mjs NEW_REPORT [chromium|firefox|webkit] [rows, 100–2000 in steps of 100]',
  );
const report = {
  engine,
  startedAt: new Date().toISOString(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  scriptSha256: createHash('sha256')
    .update(readFileSync('scripts/paste-formula-browser-benchmark.mjs'))
    .digest('hex'),
  limitations:
    'Headless single-host samples; no paired baseline, independent hardware or peak Worker/process memory measurement. Frame intervals are main-document requestAnimationFrame observations during the API paste, not input latency or total process responsiveness. Excludes native clipboard permission and operating-system input.',
  rounds: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
report.artifacts = [];
for (const file of [
  'index.html',
  ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
]) {
  const response = await fetch(new URL(file, 'http://127.0.0.1:5186/'));
  const sha256 = hash(new Uint8Array(await response.arrayBuffer()));
  if (!response.ok || sha256 !== hash(readFileSync(`apps/playground/dist/${file}`)))
    throw new Error(`Build mismatch: ${file}`);
  report.artifacts.push({ file, sha256 });
}
save();
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5186/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(async (rowCount) => {
        const app = window.opensheet;
        let book = app.getWorkbook();
        await book.generate(Math.max(100_000, rowCount * 100));
        const sheetId = book.getSheets()[0].id;
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([[1]]);
        const read = async () => {
          const values = [];
          for (let startRow = 0; startRow < rowCount; startRow += 100)
            values.push(
              ...(await book
                .getSheets()[0]
                .range({
                  startRow: startRow + 1,
                  endRow: startRow + 101,
                  startColumn: 0,
                  endColumn: 100,
                })
                .getValues()),
            );
          return values;
        };
        const original = await read();
        if (original.length !== rowCount || original.some((row) => row.length !== 100))
          throw new Error('Unexpected fixture dimensions');
        const verify = async (expected, label) => {
          const values = await read();
          for (let r = 0; r < rowCount; r++)
            for (let c = 0; c < 100; c++)
              if (values[r][c] !== expected(r, c)) throw new Error(`${label}: ${r},${c}`);
        };
        const frameIntervals = [];
        let previous = await new Promise((resolve) => requestAnimationFrame(resolve));
        let active = true;
        let frameId = 0;
        const sample = (timestamp) => {
          frameIntervals.push(timestamp - previous);
          previous = timestamp;
          if (active) frameId = requestAnimationFrame(sample);
        };
        frameId = requestAnimationFrame(sample);
        let fillMs;
        try {
          const start = performance.now();
          await book.execute({
            type: 'core.cells.pasteTile',
            payload: {
              sheetId,
              range: { startRow: 1, endRow: rowCount + 1, startColumn: 0, endColumn: 100 },
              rows: 1,
              columns: 1,
              cells: [
                { input: { type: 'formula', expression: '$A$1+1' }, source: { row: 0, column: 0 } },
              ],
            },
          });
          fillMs = performance.now() - start;
          await new Promise((resolve) => requestAnimationFrame(resolve));
        } finally {
          active = false;
          cancelAnimationFrame(frameId);
        }
        await verify(() => 2, 'paste');
        await book.undo();
        await verify((r, c) => original[r][c], 'undo');
        await book.redo();
        await verify(() => 2, 'redo');
        await app.open(book.id);
        book = app.getWorkbook();
        await verify(() => 2, 'reopen');
        return { cells: rowCount * 100, fillMs, frameIntervals };
      }, rowCount);
      report.rounds.push({ round, ...result });
      console.log(
        JSON.stringify({
          engine,
          round,
          cells: result.cells,
          fillMs: result.fillMs,
          frameMaxMs: Math.max(...result.frameIntervals),
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
  if (JSON.stringify(localBuildFingerprint()) !== JSON.stringify(report.build))
    throw new Error('Build changed during measurement');
  report.finishedAt = new Date().toISOString();
  save();
} finally {
  await browser.close();
}
