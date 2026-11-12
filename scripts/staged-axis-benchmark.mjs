import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/staged-axis-benchmark.mjs NEW_REPORT [chromium|firefox|webkit]',
  );
const initialRows = 99900,
  added = 100;
const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: initialRows, columns: 1 }] });
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  engine,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Development-server direct Workbook and real Worker; 10 staged insertions of 10 rows at the start of a 99,900-row sheet plus a write at the expanded edge. Checks every axis ID, value, undo, redo and reopen. No paired baseline, production assets, process peak memory, native input or other devices.',
  rounds: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const browser = await { chromium, firefox, webkit }[engine].launch();
try {
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5173/');
      const result = await page.evaluate(
        async ({ fixture, initialRows, added, moduleUrl }) => {
          const { Workbook, openWorkbook } = await import(/* @vite-ignore */ moduleUrl);
          let book = await new Workbook(fixture).ready(),
            sheet = book.getSheets()[0];
          const verify = async (expectedRows, withInsert) => {
            const meta = book.sheetData(sheet.id);
            if (meta.rowOrder.length !== expectedRows) throw new Error('Wrong row count');
            for (let i = 0; i < expectedRows; i++) {
              const expected =
                withInsert && i < added
                  ? initialRows + (9 - Math.floor(i / 10)) * 10 + (i % 10)
                  : i - (withInsert ? added : 0);
              if (meta.rowOrder[i] !== expected) throw new Error(`Wrong row identity ${i}`);
            }
            const value = await sheet.range(`A${expectedRows}`).getValues();
            if (value[0][0] !== (withInsert ? 42 : null)) throw new Error('Wrong edge value');
          };
          try {
            await verify(initialRows, false);
            let collectionMs = 0;
            const start = performance.now();
            await book.transaction({}, async (tx) => {
              const staged = tx.getSheets()[0],
                collect = performance.now();
              for (let i = 0; i < 10; i++) await staged.insertRows(0, 10);
              await staged.range(`A${initialRows + added}`).setValues([[42]]);
              collectionMs = performance.now() - collect;
            });
            const totalMs = performance.now() - start;
            await verify(initialRows + added, true);
            if (!(await book.undo())) throw new Error('Missing undo');
            await verify(initialRows, false);
            if (!(await book.redo())) throw new Error('Missing redo');
            await verify(initialRows + added, true);
            const id = book.id;
            await book.close();
            book = await openWorkbook(id);
            sheet = book.getSheets()[0];
            await verify(initialRows + added, true);
            return {
              collectionMs,
              totalMs,
              verifiedRows: initialRows + added,
              bytes: (await book.storageStats()).bytes,
            };
          } finally {
            await book.deleteStorage();
            book.dispose();
          }
        },
        {
          fixture: file,
          initialRows,
          added,
          moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
        },
      );
      report.rounds.push({ round, ...result });
      console.log(JSON.stringify({ round, ...result }));
    } catch (error) {
      report.rounds.push({ round, error: String(error) });
      throw error;
    } finally {
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
      await context.close();
    }
  }
} finally {
  await browser.close();
}
