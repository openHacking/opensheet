import { chromium } from '@playwright/test';
import { existsSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { cpus, platform, arch, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((value) => value.startsWith('--output='))?.slice(9);
if (!output?.endsWith('.json') || existsSync(resolve(output)))
  throw new Error('Use a new --output=NAME.json');
const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  await page.goto('http://127.0.0.1:5174/#budget');
  await page.locator('#scene-title').filter({ hasText: 'Budget Simulator' }).waitFor();
  const samples = await page.evaluate(
    async ({ coreUrl, copyUrl, encoderUrl }) => {
      const { createWorkbook, key } = await import(/* @vite-ignore */ coreUrl);
      const { copyText } = await import(/* @vite-ignore */ copyUrl);
      const { encodeClipboardText } = await import(/* @vite-ignore */ encoderUrl);
      const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 1000, columns: 10 }] });
      try {
        const sheet = book.getSheets()[0];
        const rows = Array.from({ length: 1000 }, (_, r) =>
          Array.from({ length: 10 }, (_, c) => r * 10 + c),
        );
        await sheet.range('A1:J1000').setValues(rows);
        const expected = rows.map((row) => row.join('\t')).join('\n');
        const range = sheet.range('A1:J1000');
        const baseline = async () => {
          const bounds = { ...range.bounds },
            revision = book.revision;
          const data = await book.readRange(sheet.id, bounds);
          if (data.revision !== revision || book.revision !== revision)
            throw new Error('Revision changed');
          const model = book.sheetData(sheet.id);
          const output = [];
          for (let row = bounds.startRow; row < bounds.endRow; row++) {
            const cells = [];
            for (let column = bounds.startColumn; column < bounds.endColumn; column++) {
              const address = key(model.rowOrder[row], model.columnOrder[column]);
              const value = data.calculated[address];
              if (value && typeof value === 'object' && value.error === '#LIMIT!')
                throw new Error('Incomplete calculation');
              cells.push(data.display[address] ?? '');
            }
            output.push(cells);
          }
          return encodeClipboardText(output);
        };
        const result = [];
        for (let pair = 0; pair < 3; pair++) {
          for (const variant of pair % 2 ? ['batched', 'baseline'] : ['baseline', 'batched']) {
            const start = performance.now();
            const text = await (variant === 'baseline' ? baseline() : copyText(range));
            const elapsedMs = performance.now() - start;
            if (text !== expected) throw new Error(`${variant} changed output`);
            result.push({
              pair,
              variant,
              elapsedMs,
              cells: rows.length * rows[0].length,
              outputLength: text.length,
            });
          }
        }
        return result;
      } finally {
        await book.deleteStorage();
      }
    },
    {
      coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      copyUrl: `/@fs${process.cwd()}/packages/renderer/src/clipboard.ts`,
      encoderUrl: `/@fs${process.cwd()}/packages/renderer/src/clipboard-text.ts`,
    },
  );
  const report = {
    browser: browser.version(),
    platform: platform(),
    arch: arch(),
    osRelease: release(),
    cpu: cpus()[0]?.model,
    source: sourceFingerprint(),
    fixture: '10,000 numeric cells, one 1,000 x 10 range',
    scope:
      'Alternating reconstructed prior and current copyText against the same real Worker in a Vite development server. Includes Worker reads and TSV construction; no system clipboard, process memory, production build or other hardware.',
    samples,
  };
  writeFileSync(resolve(output), JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  process.stdout.write(JSON.stringify(report, null, 2) + '\n');
} finally {
  await browser.close();
}
