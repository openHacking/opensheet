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
const boundary = `      }
      // Preconditions describe the state at the beginning of this command,
      // including when more than one entry names the same coordinate.
      for (const x of p.cells) {
        const target = cell(sh, x.row, x.column);`;
if (source.split(boundary).length !== 2) throw new Error('Cannot reconstruct the control');
const control = source.replace(boundary, '        const target = cell(sh, x.row, x.column);');

async function compile(contents) {
  const bundle = await build({
    stdin: {
      contents: `export { cellsCommand } from ${JSON.stringify(path)}; export { createSnapshot } from ${JSON.stringify(resolve('packages/core/src/model.ts'))};`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'cell-command-variant',
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
    .update(readFileSync('scripts/cell-set-precondition-benchmark.mjs'))
    .digest('hex'),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Node-only reducer, 20,000 unique cells in a fresh snapshot per sample; reconstructed one-pass control has different semantics for duplicate coordinates and is compared only on unique coordinates. Six alternating pairs for no preconditions, numeric preconditions and text preconditions with a 128-character suffix; verifies all output outside timing. No Worker, schema parsing, storage, browser UI, peak memory or independent reference device. Sequential samples in one process are sensitive to JIT and GC.',
  samples: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const rows = 200;
const columns = 100;
for (const fixture of ['none', 'numeric', 'text']) {
  for (let round = 0; round < 6; round++) {
    for (const variant of round % 2 ? ['current', 'control'] : ['control', 'current']) {
      const { cellsCommand, createSnapshot } = implementations[variant];
      const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows, columns }] });
      const sheet = snapshot.sheets[0];
      const cells = [];
      for (let row = 0; row < rows; row++) {
        for (let column = 0; column < columns; column++) {
          const input =
            fixture === 'text'
              ? { type: 'string', value: `${row}:${column}:` + 'x'.repeat(128) }
              : { type: 'number', value: row * columns + column };
          sheet.cells[`${row}:${column}`] = { rowId: row, columnId: column, input };
          cells.push({
            row,
            column,
            input: { type: 'number', value: row * columns + column + 1 },
            ...(fixture === 'none'
              ? {}
              : { expected: { rowId: row, columnId: column, input: structuredClone(input) } }),
          });
        }
      }
      const start = performance.now();
      cellsCommand(snapshot, { type: 'core.cells.set', payload: { sheetId: sheet.id, cells } });
      const elapsedMs = performance.now() - start;
      for (let row = 0; row < rows; row++)
        for (let column = 0; column < columns; column++)
          if (sheet.cells[`${row}:${column}`].input.value !== row * columns + column + 1)
            throw new Error(`Incorrect ${fixture} result at ${row}:${column}`);
      report.samples.push({ fixture, round: round + 1, variant, elapsedMs });
      save();
    }
  }
}
report.finishedAt = new Date().toISOString();
save();
console.log(JSON.stringify(report.samples));
