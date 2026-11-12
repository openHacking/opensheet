import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/fill-down-reducer-benchmark.mjs NEW_REPORT');
const cellPath = resolve('packages/core/src/commands/cells.ts');
const source = readFileSync(cellPath, 'utf8');
const current =
  'const snapshotRange = fill ? { ...source, endRow: source.startRow + 1 } : source;\n      iterate(sh, snapshotRange, (r, c) => {';
const baseline = 'iterate(sh, source, (r, c) => {';
if (source.split(current).length !== 2) throw new Error('Cannot reconstruct baseline safely');

async function compile(code) {
  const result = await build({
    stdin: {
      contents: `export { cellsCommand } from ${JSON.stringify(cellPath)}; export { createSnapshot } from ${JSON.stringify(resolve('packages/core/src/model.ts'))};`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'variant',
        setup: (builder) =>
          builder.onLoad({ filter: /commands\/cells\.ts$/ }, () => ({
            contents: code,
            loader: 'ts',
            resolveDir: resolve('packages/core/src/commands'),
          })),
      },
    ],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
  });
  return import(
    `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].contents).toString('base64')}`
  );
}

const implementations = {
  baseline: await compile(source.replace(current, baseline)),
  current: await compile(source),
};
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  rows: 10000,
  columns: 10,
  limitations:
    'Node reducer only; excludes Worker RPC, storage, undo and browser memory. Alternating variants in one process; no isolated peak-memory measurement.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 6; round++) {
  for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
    const { cellsCommand, createSnapshot } = implementations[variant];
    const snapshot = createSnapshot({
      sheets: [{ name: 'Sheet', rows: report.rows, columns: report.columns }],
    });
    const sheet = snapshot.sheets[0];
    for (let row = 0; row < report.rows; row++)
      for (let column = 0; column < report.columns; column++)
        sheet.cells[`${row}:${column}`] = {
          rowId: row,
          columnId: column,
          input: { type: 'number', value: row + column },
        };
    const command = {
      type: 'core.cells.fillDown',
      payload: {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: report.rows, startColumn: 0, endColumn: report.columns },
      },
    };
    const start = performance.now();
    cellsCommand(snapshot, command);
    const elapsedMs = performance.now() - start;
    for (let row = 0; row < report.rows; row++)
      for (let column = 0; column < report.columns; column++)
        if (sheet.cells[`${row}:${column}`].input.value !== column)
          throw new Error('Incorrect fill result');
    report.samples.push({ round: round + 1, variant, elapsedMs });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report.samples.at(-1)));
  }
}
