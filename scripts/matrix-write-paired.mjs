import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Provide a new report path');
const path = resolve('packages/core/src/range.ts');
const source = readFileSync(path, 'utf8');
const start = source.indexOf('  private write<T>(');
const end = source.indexOf('  setValues(values:', start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct control');
const control =
  source.slice(0, start) +
  `  private write<T>(matrix: T[][], map: (value: T) => CellInput) {
    const b = this.bounds;
    assert(matrix.length === b.endRow - b.startRow && matrix.every((row) => row.length === b.endColumn - b.startColumn), 'INVALID_ARGUMENT', 'Matrix dimensions must match the range');
    return this.sheet.workbook.execute({type:'core.cells.set',payload:{sheetId:this.sheet.id,cells:matrix.flatMap((row,i)=>row.map((value,j)=>({row:b.startRow+i,column:b.startColumn+j,input:map(value)})))}});
  }
` +
  source.slice(end);
async function compile(contents) {
  const bundle = await build({
    stdin: {
      contents: `export { Range } from ${JSON.stringify(path)};`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    plugins: [
      {
        name: 'range-variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/range\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/core/src'),
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
const variants = { control: await compile(control), current: await compile(source) };
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  scriptSha256: createHash('sha256')
    .update(readFileSync('scripts/matrix-write-paired.mjs'))
    .digest('hex'),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'Node-only Range matrix-to-command construction; 100,000 dense numeric or formula inputs, six alternating fresh-range pairs. Control reconstructs the prior flatMap implementation and silently skips sparse entries. Every generated input and coordinate is verified outside timing. Excludes Worker transfer, schema validation, persistence, UI, peak memory and independent hardware; sequential samples share JIT and GC.',
  samples: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
const rows = 1000,
  columns = 100;
for (const fixture of ['numeric', 'formula']) {
  const matrix = Array.from({ length: rows }, (_, r) =>
    Array.from({ length: columns }, (_, c) => (fixture === 'numeric' ? r * columns + c : '=A1+1')),
  );
  for (let round = 0; round < 6; round++)
    for (const variant of round % 2 ? ['current', 'control'] : ['control', 'current']) {
      let command;
      const sheet = {
        id: 'sheet',
        workbook: {
          execute(value) {
            command = value;
          },
        },
      };
      const range = new variants[variant].Range(sheet, {
        startRow: 0,
        endRow: rows,
        startColumn: 0,
        endColumn: columns,
      });
      const startTime = performance.now();
      if (fixture === 'numeric') range.setValues(matrix);
      else range.setFormulas(matrix);
      const elapsedMs = performance.now() - startTime;
      if (command.payload.cells.length !== rows * columns) throw new Error('Wrong count');
      for (let i = 0; i < rows * columns; i++) {
        const cell = command.payload.cells[i];
        if (
          cell.row !== Math.floor(i / columns) ||
          cell.column !== i % columns ||
          (fixture === 'numeric'
            ? cell.input.type !== 'number' || cell.input.value !== i
            : cell.input.type !== 'formula' || cell.input.expression !== 'A1+1')
        )
          throw new Error(`Incorrect input ${fixture}/${i}`);
      }
      report.samples.push({ fixture, round: round + 1, variant, elapsedMs });
      save();
    }
}
report.finishedAt = new Date().toISOString();
save();
console.log(JSON.stringify(report.samples));
