import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/structural-delete-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const rows = 512,
  columns = 16;
const file = createWorkbookFile({
  sheets: [
    { name: 'Inputs', rows: 256, columns: 1 },
    { name: 'Formulas', rows, columns },
  ],
});
file.sheets[0].blocks.push({
  row: 0,
  column: 0,
  data: { numbers: { positions: [0, 32], values: [41, 7] } },
});
for (let page = 0; page < rows / 64; page++) {
  const start = page * 64 * columns;
  file.sheets[1].blocks.push({
    row: page,
    column: 0,
    data: {
      formulas: {
        positions: Array.from(
          { length: 64 * columns },
          (_, i) => Math.floor(i / columns) * 32 + (i % columns),
        ),
        values: Array.from(
          { length: 64 * columns },
          (_, i) => `SUM(Inputs!A1:A2)+${start + i + 1}`,
        ),
      },
    },
  });
}
const report = {
  engine,
  cpu: cpus()[0]?.model,
  os: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  startedAt: new Date().toISOString(),
  formulaCells: rows * columns,
  scope:
    'Development server; real Worker/IndexedDB deletion of first row partially overlaps every formula range. Every result checked before, after, undo, redo and reopen. No paired baseline, process peak memory, native input or reference devices.',
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
      await page.goto('http://127.0.0.1:5173/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(
        async ({ fixture, rows, columns }) => {
          const app = window.opensheet;
          await app.load(fixture);
          let book = app.getWorkbook();
          const check = async (base) => {
            const values = await book
              .getSheets()[1]
              .range({ startRow: 0, endRow: rows, startColumn: 0, endColumn: columns })
              .getValues();
            for (let row = 0; row < rows; row++)
              for (let col = 0; col < columns; col++) {
                const expected = base + row * columns + col + 1;
                if (values[row][col] !== expected)
                  throw new Error(
                    `Mismatch at ${row},${col}: ${JSON.stringify(values[row][col])} != ${expected}`,
                  );
              }
          };
          await check(48);
          const start = performance.now();
          await book.getSheets()[0].deleteRows(0);
          const deleteMs = performance.now() - start;
          await check(7);
          if (!(await book.undo())) throw new Error('Missing undo');
          await check(48);
          if (!(await book.redo())) throw new Error('Missing redo');
          await check(7);
          await app.open(book.id);
          book = app.getWorkbook();
          await check(7);
          return {
            deleteMs,
            verifiedPerCheck: rows * columns,
            encodedBytes: (await book.storageStats()).bytes,
          };
        },
        { fixture: file, rows, columns },
      );
      report.rounds.push({ round, ...result });
      console.log(JSON.stringify({ round, ...result }));
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
