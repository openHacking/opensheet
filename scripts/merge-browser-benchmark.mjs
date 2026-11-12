import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/merge-browser-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const rows = 20000;
const file = createWorkbookFile({ sheets: [{ name: 'Data', rows, columns: 1 }] });
for (let page = 0; page < Math.ceil(rows / 64); page++) {
  const count = Math.min(64, rows - page * 64),
    positions = Array.from({ length: count }, (_, i) => i * 32);
  file.sheets[0].blocks.push({
    row: page,
    column: 0,
    data: {
      blanks: positions,
      notes: { positions, values: positions.map((_, i) => `note-${page * 64 + i}`) },
    },
  });
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  engine,
  cells: rows,
  scope:
    'Production preview; real Worker and IndexedDB merge of 20,000 existing annotated blank cells. All notes verified before, after, undo, redo and reopen. No paired baseline, process peak memory, native input or other devices.',
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
      const result = await page.evaluate(
        async ({ fixture, rows }) => {
          const app = window.opensheet;
          await app.load(fixture);
          let book = app.getWorkbook(),
            sheet = book.getSheets()[0];
          const range = { startRow: 0, endRow: rows, startColumn: 0, endColumn: 1 };
          const check = async (expected) => {
            for (let start = 0; start < rows; start += 8192) {
              const end = Math.min(rows, start + 8192);
              const notes = await sheet
                .range({ ...range, startRow: start, endRow: end })
                .getNotes();
              if (notes.length !== end - start) throw new Error('Incorrect row count');
              for (let row = start; row < end; row++)
                if (notes[row - start][0] !== `note-${row}`) throw new Error(`Changed note ${row}`);
            }
            if (book.sheetData(sheet.id).merges.length !== expected)
              throw new Error('Incorrect merge state');
          };
          await check(0);
          const started = performance.now();
          await sheet.range(range).merge();
          const mergeMs = performance.now() - started;
          await check(1);
          if (!(await book.undo())) throw new Error('Missing undo');
          await check(0);
          if (!(await book.redo())) throw new Error('Missing redo');
          await check(1);
          await app.open(book.id);
          book = app.getWorkbook();
          sheet = book.getSheets()[0];
          await check(1);
          return { mergeMs, verifiedNotesPerCheck: rows, bytes: (await book.storageStats()).bytes };
        },
        { fixture: file, rows },
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
