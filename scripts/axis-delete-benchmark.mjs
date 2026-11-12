import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new JSON report path');
const path = resolve('packages/core/src/commands/axis.ts');
const source = readFileSync(path, 'utf8');
const needle = 'const removed = new Set(deleted);';
const lookup = 'removed.has(row ? sh.cells[k].rowId : sh.cells[k].columnId)';
if (source.split(needle).length !== 2 || source.split(lookup).length !== 2)
  throw new Error('Cannot reconstruct control');
const control = source
  .replace(needle, '')
  .replace(lookup, 'deleted.includes(row ? sh.cells[k].rowId : sh.cells[k].columnId)');
async function compile(contents) {
  const bundle = await build({
    stdin: {
      contents: `export { axisCommand } from ${JSON.stringify(path)}; export { createSnapshot } from ${JSON.stringify(resolve('packages/core/src/model.ts'))};`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'axis-variant',
        setup(builder) {
          builder.onLoad({ filter: /commands\/axis\.ts$/ }, () => ({
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
    .update(readFileSync('scripts/axis-delete-benchmark.mjs'))
    .digest('hex'),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Node reducer only; delete 19,999 of 20,000 populated rows, no formulas or merges; six alternating pairs with fresh snapshots. Control reconstructs array membership instead of Set membership; all remaining cells, order and metadata verified outside timing. Single process, JIT and GC variation; excludes Worker, persistence, UI, peak memory and independent devices.',
  samples: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (let round = 0; round < 6; round++) {
  for (const variant of round % 2 ? ['current', 'control'] : ['control', 'current']) {
    const { axisCommand, createSnapshot } = implementations[variant];
    const snapshot = createSnapshot({ sheets: [{ name: 'Sheet', rows: 20000, columns: 1 }] });
    const sheet = snapshot.sheets[0];
    for (let row = 0; row < 20000; row++)
      sheet.cells[`${row}:0`] = { rowId: row, columnId: 0, input: { type: 'number', value: row } };
    sheet.rows[19999] = { size: 30 };
    const started = performance.now();
    axisCommand(snapshot, {
      type: 'core.axis.delete',
      payload: { sheetId: sheet.id, axis: 'row', index: 0, count: 19999 },
    });
    const elapsedMs = performance.now() - started;
    const remaining = Object.values(sheet.cells);
    if (
      remaining.length !== 1 ||
      remaining[0].rowId !== 19999 ||
      remaining[0].input.value !== 19999 ||
      sheet.rowOrder.length !== 1 ||
      sheet.rowOrder[0] !== 19999 ||
      sheet.rows[19999]?.size !== 30
    )
      throw new Error('Incorrect deletion');
    report.samples.push({ round: round + 1, variant, elapsedMs });
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
console.log(JSON.stringify(report.samples));
