import { chromium } from '@playwright/test';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { writeFileSync } from 'node:fs';
import { cpus, totalmem, platform, arch, release } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Pass a new output path');
const rows = 100000;
const cooperative = process.argv.includes('--cooperative');
const browser = await chromium.launch();
const report = {
  startedAt: new Date().toISOString(),
  rows,
  cooperative,
  sourceFingerprint: sourceFingerprint(),
  localBuildFingerprint: localBuildFingerprint(),
  browser: browser.version(),
  node: process.version,
  hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length, memoryBytes: totalmem() },
  platform: { platform: platform(), arch: arch(), release: release() },
  powerMode: 'unrecorded',
  scope:
    'Headless Chromium, local development server; complete Worker search RPC time, no Worker memory or cross-device claim.',
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
        async ({ rows, file, cooperative }) => {
          const app = window.opensheet;
          await app.load(file);
          const book = app.getWorkbook();
          const sheet = book.getSheets()[0];
          const revision = book.revision;
          let limitCode = '';
          try {
            await sheet
              .range(`A1:A${rows}`)
              .setValues(Array.from({ length: rows }, () => ['probe']));
          } catch (error) {
            limitCode = error?.code ?? '';
          }
          if (
            limitCode !== 'LIMIT_EXCEEDED' ||
            book.revision !== revision ||
            (await sheet.range('A1').getValues())[0][0] !== null
          )
            throw new Error(`Unexpected oversized-write result: ${limitCode}`);
          for (let start = 0; start < rows; start += 10000) {
            const end = Math.min(start + 10000, rows);
            await sheet
              .range(`A${start + 1}:A${end}`)
              .setValues(
                Array.from({ length: end - start }, (_, i) => [
                  `entry-${String(start + i).padStart(6, '0')}`,
                ]),
              );
          }
          const samples = [];
          for (const [query, after, expectedCount, expectedRow] of [
            ['entry-', -1, rows, 0],
            ['entry-', 0, rows, 1],
            ['entry-099999', -1, 1, rows - 1],
            ['missing', -1, 0, -1],
          ]) {
            const start = performance.now();
            const signal = cooperative ? new AbortController().signal : undefined;
            const result = await book.search(sheet.id, query, after, { signal });
            const elapsedMs = performance.now() - start;
            if (
              result.count !== expectedCount ||
              (result.match?.row ?? -1) !== expectedRow ||
              (result.match && result.match.column !== 0)
            )
              throw new Error(`Unexpected search result for ${query}: ${JSON.stringify(result)}`);
            samples.push({
              query,
              after,
              elapsedMs,
              count: result.count,
              match: result.match ?? null,
            });
          }
          const id = book.id;
          await app.open(id);
          const reopened = await app.getWorkbook().search(sheet.id, 'entry-099999');
          if (reopened.count !== 1 || reopened.match?.row !== rows - 1)
            throw new Error('Reopen search mismatch');
          return { limitCode, samples, reopened };
        },
        {
          rows,
          cooperative,
          file: createWorkbookFile({ sheets: [{ name: 'Search', rows, columns: 26 }] }),
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
