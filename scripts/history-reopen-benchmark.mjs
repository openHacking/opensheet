import { chromium, firefox, webkit } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { cpus, totalmem, platform, arch, release } from 'node:os';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium', workload = 'cells'] = process.argv.slice(2);
if (
  !output ||
  !['chromium', 'firefox', 'webkit'].includes(engine) ||
  !['cells', 'metadata'].includes(workload)
)
  throw new Error(
    'Usage: node scripts/history-reopen-benchmark.mjs NEW_REPORT [chromium|firefox|webkit] [cells|metadata]',
  );
const browser = await { chromium, firefox, webkit }[engine].launch();
const report = {
  startedAt: new Date().toISOString(),
  engine,
  workload,
  browser: browser.version(),
  node: process.version,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length, memoryBytes: totalmem() },
  platform: { platform: platform(), arch: arch(), release: release() },
  powerMode: 'unrecorded',
  scope: `Headless local production preview; three fresh contexts, 100 ${workload === 'metadata' ? 'sheet renames with distinct stored metadata' : 'multi-page edits'} and complete undo/redo checks. Reopen includes editor attachment and first viewport, not just history validation. No Worker heap or cross-device claim.`,
  rounds: [],
};
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
const file = createWorkbookFile({
  sheets: [{ name: 'History', rows: workload === 'metadata' ? 8192 : 256, columns: 32 }],
});
if (workload === 'metadata') {
  file.sheets[0].rows.meta = Object.fromEntries(
    Array.from({ length: 512 }, (_, i) => [String(i * 16), { size: 30 + (i % 5) }]),
  );
  file.sheets[0].blocks.push({
    row: 0,
    column: 0,
    data: { numbers: { positions: [0], values: [5] } },
  });
}
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  for (let round = 1; round <= 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5186/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(
        async ({ file, workload }) => {
          const app = window.opensheet;
          await app.load(file);
          let book = app.getWorkbook();
          let sheet = book.getSheets()[0];
          const expected = (step, row, column) =>
            step === 0 ? null : step * 100000 + row * 32 + column;
          const start = performance.now();
          for (let step = 1; step <= 100; step++) {
            if (workload === 'metadata') {
              await book.renameSheet(sheet.id, `History ${step}`);
              continue;
            }
            await sheet
              .range('A1:AF256')
              .setValues(
                Array.from({ length: 256 }, (_, row) =>
                  Array.from({ length: 32 }, (_, column) => expected(step, row, column)),
                ),
              );
          }
          const setupMs = performance.now() - start;
          const id = book.id;
          const reopenMs = [];
          const verifyMetadata = async (step) => {
            if (
              sheet.name !== (step ? `History ${step}` : 'History') ||
              (await sheet.range('A1').getValues())[0][0] !== 5
            )
              throw new Error(`Metadata history mismatch at step ${step}`);
          };
          const verifyAll = async (step) => {
            if (workload === 'metadata') return verifyMetadata(step);
            const cells = await sheet.range('A1:AF256').getValues();
            for (let row = 0; row < 256; row++)
              for (let column = 0; column < 32; column++)
                if (cells[row][column] !== expected(step, row, column))
                  throw new Error(`Mismatch ${step}/${row}/${column}`);
          };
          const verifyEdges = async (step) => {
            if (workload === 'metadata') return verifyMetadata(step);
            const values = await Promise.all([
              sheet.range('A1').getValues(),
              sheet.range('AF256').getValues(),
            ]);
            if (
              values[0][0][0] !== expected(step, 0, 0) ||
              values[1][0][0] !== expected(step, 255, 31)
            )
              throw new Error(`History mismatch at step ${step}`);
          };
          for (let i = 0; i < 3; i++) {
            const before = performance.now();
            await app.open(id);
            reopenMs.push(performance.now() - before);
            book = app.getWorkbook();
            sheet = book.getSheets()[0];
            await verifyAll(100);
          }
          const undoStart = performance.now();
          for (let step = 99; step >= 0; step--) {
            if (!(await book.undo())) throw new Error(`Missing undo ${step}`);
            await verifyEdges(step);
          }
          const undoMs = performance.now() - undoStart;
          await verifyAll(0);
          const redoStart = performance.now();
          for (let step = 1; step <= 100; step++) {
            if (!(await book.redo())) throw new Error(`Missing redo ${step}`);
            await verifyEdges(step);
          }
          const redoMs = performance.now() - redoStart;
          await verifyAll(100);
          const stats = await book.storageStats();
          await book.deleteStorage();
          return {
            setupMs,
            reopenMs,
            undoMs,
            redoMs,
            verifiedCellsPerFullCheck: workload === 'metadata' ? 1 : 8192,
            undoSteps: 100,
            redoSteps: 100,
            stats,
          };
        },
        { file, workload },
      );
      report.rounds.push({ round, ...result });
      save();
      process.stdout.write(
        JSON.stringify({
          engine,
          round,
          reopenMs: result.reopenMs,
          undoMs: result.undoMs,
          redoMs: result.redoMs,
        }) + '\n',
      );
    } catch (error) {
      report.rounds.push({ round, error: String(error) });
      save();
      throw error;
    } finally {
      await context.close();
    }
  }
} finally {
  await browser.close();
}
