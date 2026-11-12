import { chromium } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync } from 'node:fs';
import { cpus, totalmem, platform, arch, release } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Pass a new output path');
const rows = 10000;
const browser = await chromium.launch();
const report = {
  startedAt: new Date().toISOString(),
  rows,
  sourceFingerprint: sourceFingerprint(),
  localBuildFingerprint: localBuildFingerprint(),
  browser: browser.version(),
  node: process.version,
  hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length, memoryBytes: totalmem() },
  platform: { platform: platform(), arch: arch(), release: release() },
  powerMode: 'unrecorded',
  scope:
    'Headless Chromium on a local development server. Separate acknowledged edit and recalculation/viewport read timings; no Worker heap or reference-device claim.',
  rounds: [],
};
writeFileSync(output, JSON.stringify(report, null, 2), { flag: 'wx' });
try {
  for (let round = 0; round < 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto('http://127.0.0.1:5173/#budget');
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(
        async ({ rows, file }) => {
          const app = window.opensheet;
          await app.load(file);
          const book = app.getWorkbook();
          const sheet = book.getSheets()[0];
          await sheet
            .range(`A1:A${rows}`)
            .setValues(Array.from({ length: rows }, (_, i) => [i + 1]));
          await sheet.range('B1').setValues([[1]]);
          await sheet.range('C1').setFormulas([[`MATCH(B1,A1:A${rows},1)`]]);
          const samples = [];
          for (let i = 0; i < 10; i++) {
            const target = 250 + i * 100;
            let start = performance.now();
            await sheet.range('B1').setValues([[target]]);
            const editMs = performance.now() - start;
            start = performance.now();
            const value = (await sheet.range('C1').getValues())[0][0];
            const calculationMs = performance.now() - start;
            if (value !== target) throw new Error(`Unexpected MATCH result: ${value}`);
            samples.push({ target, editMs, calculationMs });
          }
          await book.undo();
          const undone = (await sheet.range('C1').getValues())[0][0];
          await book.redo();
          const redone = (await sheet.range('C1').getValues())[0][0];
          if (undone !== 1050 || redone !== 1150) throw new Error('Undo/redo mismatch');
          const id = book.id;
          await app.open(id);
          const reopened = (await app.getWorkbook().getSheets()[0].range('C1').getValues())[0][0];
          if (reopened !== 1150) throw new Error('Reopen mismatch');
          return { samples, undone, redone, reopened };
        },
        {
          rows,
          file: createWorkbookFile({ sheets: [{ name: 'Ordered lookup', rows, columns: 26 }] }),
        },
      );
      report.rounds.push({ round: round + 1, ...result });
    } catch (error) {
      report.rounds.push({ round: round + 1, error: String(error) });
      throw error;
    } finally {
      await context.close();
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    }
  }
} finally {
  await browser.close();
}
