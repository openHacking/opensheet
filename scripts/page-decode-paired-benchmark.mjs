import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/page-decode-paired-benchmark.mjs NEW_REPORT');
const blockPath = resolve('packages/core/src/block-codec.ts');
const filePath = resolve('packages/core/src/file-codec.ts');
const currentBlock = readFileSync(blockPath, 'utf8');
const currentFile = readFileSync(filePath, 'utf8');
const call = '      undefined,\n      undefined,\n      inlineKinds.size === 0,\n';
if (currentBlock.split(call).length !== 2) throw new Error('Cannot reconstruct previous page call');
const start = currentFile.indexOf(
  '  const result: Record<string, CellRecord> = {};\n',
  currentFile.indexOf('export function decodeFileCells('),
);
const end = currentFile.indexOf('\n}\nexport function encodeWorkbookFile', start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct previous cell table');
const previousFile =
  currentFile.slice(0, start) +
  '  return Object.fromEntries(\n    [...cells.values()].map((cell) => [key(cell.rowId, cell.columnId), cell]),\n  );' +
  currentFile.slice(end);
const scripts = {},
  hashes = {};
for (const variant of ['previous', 'current']) {
  const bundle = await build({
    entryPoints: [blockPath],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: `Page_${variant}`,
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/(block-codec|file-codec)\.ts$/ }, (args) => ({
            contents:
              args.path === blockPath
                ? variant === 'previous'
                  ? currentBlock.replace(call, '')
                  : currentBlock
                : variant === 'previous'
                  ? previousFile
                  : currentFile,
            loader: 'ts',
            resolveDir: resolve('packages/core/src'),
          }));
        },
      },
    ],
  });
  scripts[variant] = bundle.outputFiles[0].text;
  hashes[variant] = createHash('sha256').update(scripts[variant]).digest('hex');
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Six alternating paired rounds per headless browser; 16 mixed scalar pages of 2048 cells each, decode only. All decoded values and counts are checked after timing. Previous reconstructs the full temporary cell table, not a prior commit. Excludes Worker, IndexedDB, paint, metadata/formulas, process memory, other devices and end-to-end latency.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    for (const script of Object.values(scripts)) await page.addScriptTag({ content: script });
    for (let round = 0; round < 6; round++) {
      result.rounds.push(
        await page.evaluate((round) => {
          const fixture = [];
          for (let p = 0; p < 16; p++) {
            const cells = {};
            for (let i = 0; i < 2048; i++) {
              const row = p * 64 + Math.floor(i / 32),
                column = i % 32;
              const value = i % 10 < 6 ? i + p / 10 : i % 10 < 9 ? `v-${p}-${i}` : !!(i % 20);
              const input = {
                type:
                  typeof value === 'number'
                    ? 'number'
                    : typeof value === 'boolean'
                      ? 'boolean'
                      : 'string',
                value,
              };
              cells[`${row}:${column}`] = { rowId: row, columnId: column, input };
            }
            fixture.push({ cells, bytes: globalThis.Page_current.encodeBlock(cells, [], [], {}) });
          }
          const samples = [];
          for (const variant of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
            const codec = globalThis[`Page_${variant}`];
            const start = performance.now();
            const pages = fixture.map(({ bytes }) => codec.decodeBlock(bytes));
            const ms = performance.now() - start;
            for (let p = 0; p < pages.length; p++) {
              const records = [...pages[p].records()];
              if (records.length !== 2048) throw new Error('Wrong cell count');
              for (const cell of records) {
                const expected = fixture[p].cells[`${cell.rowId}:${cell.columnId}`];
                if (
                  !expected ||
                  cell.input.type !== expected.input.type ||
                  cell.input.value !== expected.input.value
                )
                  throw new Error('Wrong decoded value');
              }
            }
            samples.push({ variant, ms, pages: pages.length, cells: pages.length * 2048 });
          }
          return samples;
        }, round),
      );
      save();
    }
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
