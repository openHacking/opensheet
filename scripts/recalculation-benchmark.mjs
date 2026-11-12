import { chromium, firefox, webkit } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync } from 'node:fs';
import { cpus, totalmem, platform, arch, release } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const [output, engine = 'chromium', url = 'http://127.0.0.1:5173/'] = process.argv.slice(2);
if (!output || !['chromium', 'firefox', 'webkit'].includes(engine))
  throw new Error(
    'Usage: node scripts/recalculation-benchmark.mjs NEW_REPORT [chromium|firefox|webkit] [URL]',
  );
const target = new URL(url);
target.hash = 'budget';
const rows = 128;
const browser = await { chromium, firefox, webkit }[engine].launch();
const report = {
  startedAt: new Date().toISOString(),
  engine,
  rows,
  sourceFingerprint: sourceFingerprint(),
  localBuildFingerprint: localBuildFingerprint(),
  browser: browser.version(),
  node: process.version,
  hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length, memoryBytes: totalmem() },
  platform: { platform: platform(), arch: arch(), release: release() },
  powerMode: 'unrecorded',
  url: target.href,
  scope:
    'Headless browser, served application; acknowledged edits and full output reads are separate. Served asset identity is not independently verified. No worker memory, internal phase attribution or cross-device acceptance.',
  rounds: [],
};
writeFileSync(output, JSON.stringify(report, null, 2), { flag: 'wx' });
try {
  for (let round = 0; round < 3; round++) {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      await page.goto(target.href);
      await page.waitForFunction(() => !!window.opensheet?.selection.get());
      const result = await page.evaluate(
        async ({ file, rows }) => {
          const app = window.opensheet;
          await app.load(file);
          const book = app.getWorkbook();
          const [source, fan, chain, overlap] = book.getSheets();
          await source.range('A1:A32').setValues(Array.from({ length: 32 }, () => [1]));
          await fan
            .range(`A1:B${rows}`)
            .setFormulas(Array.from({ length: rows }, (_, i) => [`Source!$A$1+${i}`, `${i}+1000`]));
          await chain
            .range(`A1:A${rows}`)
            .setFormulas(Array.from({ length: rows }, (_, i) => [i ? `A${i}+1` : 'Source!$A$1']));
          await overlap
            .range(`A1:A${rows}`)
            .setFormulas(Array.from({ length: rows }, (_, i) => [`SUM(Source!$A$1:$A$32)+${i}`]));
          const verify = async (
            target,
            [fanSheet, chainSheet, overlapSheet] = [fan, chain, overlap],
          ) => {
            const outputs = [];
            for (const [name, sheet, range] of [
              ['fanout', fanSheet, `A1:B${rows}`],
              ['chain', chainSheet, `A1:A${rows}`],
              ['overlap', overlapSheet, `A1:A${rows}`],
            ]) {
              const start = performance.now();
              const values = await sheet.range(range).getValues();
              outputs.push({ name, readMs: performance.now() - start });
              for (let i = 0; i < rows; i++) {
                const expected = name === 'overlap' ? target + 31 + i : target + i;
                if (values[i][0] !== expected || (name === 'fanout' && values[i][1] !== i + 1000))
                  throw new Error(
                    `${name} row ${i + 1}: ${JSON.stringify(values[i])}, expected ${expected}`,
                  );
              }
            }
            return outputs;
          };
          await verify(1);
          const samples = [];
          for (let target = 2; target <= 6; target++) {
            const start = performance.now();
            await source.range('A1').setValues([[target]]);
            const editMs = performance.now() - start;
            samples.push({ target, editMs, outputs: await verify(target) });
          }
          await book.undo();
          await verify(5);
          await book.redo();
          await verify(6);
          const id = book.id;
          await app.open(id);
          const reopened = app.getWorkbook();
          await verify(6, reopened.getSheets().slice(1));
          return {
            samples,
            verifiedOutputsPerSample: rows * 4,
            undoRedo: true,
            reopenedSamples: rows * 4,
          };
        },
        {
          rows,
          file: createWorkbookFile({
            sheets: ['Source', 'Fan', 'Chain', 'Overlap'].map((name) => ({
              name,
              rows,
              columns: 2,
            })),
          }),
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
