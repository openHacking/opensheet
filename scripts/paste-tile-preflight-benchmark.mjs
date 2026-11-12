import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const path = resolve('packages/core/src/commands/cells.ts');
const source = readFileSync(path, 'utf8');
const start = source.indexOf('      // Relocation can grow a formula');
const end = source.indexOf('      break;', start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct control');
const control =
  source.slice(0, start) +
  `      iterate(sh, tile.range, (row, column) => {
        applyPastedCell(s, cell(sh, row, column), tileCell(tile, row, column, charge, offset));
      });
` +
  source.slice(end);

async function compile(contents) {
  const bundle = await build({
    stdin: {
      contents: `export { cellsCommand } from ${JSON.stringify(path)}; export { createSnapshot } from ${JSON.stringify(resolve('packages/core/src/model.ts'))}; export { offsetFormula } from '@opensheetjs/formula';`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /commands\/cells\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/core/src/commands'),
          }));
        },
      },
    ],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
  });
  return import(
    `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
  );
}
const implementations = { control: await compile(control), current: await compile(source) };
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  scriptSha256: createHash('sha256')
    .update(readFileSync('scripts/paste-tile-preflight-benchmark.mjs'))
    .digest('hex'),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Node reducer only; 20,000 cells, numeric and repeated relative-formula tiles; six alternating pairs per fixture with fresh snapshots and all outputs verified outside timing. Control reconstructs the previous one-pass loop, which does not preserve snapshot atomicity on late failure. Single process, JIT and GC variation; excludes Worker, schema validation, storage, UI, peak memory and independent hardware.',
  samples: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const rows = 200;
const columns = 100;
const formula = 'A1+$B$1';
for (const fixture of ['numeric', 'formula']) {
  for (let round = 0; round < 6; round++) {
    for (const variant of round % 2 ? ['current', 'control'] : ['control', 'current']) {
      const { cellsCommand, createSnapshot, offsetFormula } = implementations[variant];
      const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows, columns }] });
      const sheet = snapshot.sheets[0];
      const cells =
        fixture === 'numeric'
          ? [{ input: { type: 'number', value: 42 } }]
          : [{ input: { type: 'formula', expression: formula }, source: { row: 0, column: 0 } }];
      const payload = {
        sheetId: sheet.id,
        range: { startRow: 0, endRow: rows, startColumn: 0, endColumn: columns },
        rows: 1,
        columns: 1,
        cells,
      };
      const start = performance.now();
      cellsCommand(snapshot, { type: 'core.cells.pasteTile', payload });
      const elapsedMs = performance.now() - start;
      for (let row = 0; row < rows; row++)
        for (let column = 0; column < columns; column++) {
          const input = sheet.cells[`${row}:${column}`]?.input;
          if (
            fixture === 'numeric'
              ? input?.type !== 'number' || input.value !== 42
              : input?.type !== 'formula' ||
                input.expression !== offsetFormula(formula, row, column)
          )
            throw new Error(`Incorrect ${fixture} at ${row}:${column}`);
        }
      report.samples.push({ fixture, round: round + 1, variant, elapsedMs });
      save();
    }
  }
}
report.finishedAt = new Date().toISOString();
save();
console.log(JSON.stringify(report.samples));
