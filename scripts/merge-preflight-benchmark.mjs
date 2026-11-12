import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/merge-preflight-benchmark.mjs NEW_REPORT');
const source = readFileSync('packages/core/src/commands/cells.ts', 'utf8');
const current = `        if (x) covered.push(x);
      });
      for (const x of covered) {
        x.input = { type: 'blank' };
        delete x.cached;
      }
      sh.merges.push(p.range);`;
const old = `        if (x) x.input = { type: 'blank' };
      });
      sh.merges.push(p.range);`;
if (source.split(current).length !== 2)
  throw new Error('Cannot reconstruct prior merge implementation');
async function compile(code) {
  const bundle = await build({
    stdin: {
      contents:
        "export { createSnapshot } from './packages/core/src/model.ts'; export { reduceCommand } from './packages/core/src/commands.ts';",
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'merge-variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/commands\/cells\.ts$/ }, () => ({
            contents: code,
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
const variants = {
  before: await compile(
    source.replace('      const covered: CellRecord[] = [];\n', '').replace(current, old),
  ),
  current: await compile(source),
};
const rows = 20000,
  report = {
    startedAt: new Date().toISOString(),
    source: sourceFingerprint(),
    node: process.version,
    cpu: cpus()[0]?.model,
    platform: `${process.platform}/${process.arch}/${release()}`,
    cells: rows,
    scope:
      'Alternating local Node snapshot reducer merges of 20,000 blank annotated cells. Control reconstructs the previous one-pass implementation; input clone and output verification are outside timing. No Worker, IndexedDB, browser, user input or peak memory.',
    samples: [],
  };
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const initial = variants.current.createSnapshot({ sheets: [{ name: 'Data', rows, columns: 1 }] });
const original = initial.sheets[0];
for (let row = 0; row < rows; row++)
  original.cells[`${row}:0`] = { rowId: row, columnId: 0, input: { type: 'blank' }, note: 'x' };
const range = { startRow: 0, endRow: rows, startColumn: 0, endColumn: 1 };
for (let round = 0; round < 4; round++)
  for (const variant of round % 2 ? ['current', 'before'] : ['before', 'current']) {
    const snapshot = structuredClone(initial),
      sheet = snapshot.sheets[0];
    const start = performance.now();
    variants[variant].reduceCommand(snapshot, {
      type: 'core.cells.merge',
      payload: { sheetId: sheet.id, range, discardCoveredValues: false },
    });
    const elapsedMs = performance.now() - start;
    if (sheet.merges.length !== 1 || Object.keys(sheet.cells).length !== rows)
      throw new Error('Invalid merge');
    for (let row = 0; row < rows; row++)
      if (sheet.cells[`${row}:0`].input.type !== 'blank' || sheet.cells[`${row}:0`].note !== 'x')
        throw new Error(`Cell ${row} changed`);
    report.samples.push({ round: round + 1, variant, elapsedMs, verified: rows });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  }
console.log(JSON.stringify(report.samples));
