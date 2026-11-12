import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
const distinct = process.argv[3] === 'distinct';
if (!output)
  throw new Error('Usage: node scripts/precedence-import-benchmark.mjs NEW_REPORT [distinct]');
const path = resolve('packages/adapter-sheetjs/src/cells.ts');
const current = readFileSync(path, 'utf8');
const target = "const normalized = normalize(formula, 'spreadsheet');";
if (current.split(target).length !== 2) throw new Error('Cannot construct the no-rewrite baseline');
const sources = { baseline: current.replace(target, 'const normalized = formula;'), current };
const variants = {},
  hashes = {};
for (const [name, contents] of Object.entries(sources)) {
  const result = await build({
    entryPoints: [resolve('packages/adapter-sheetjs/src/index.ts')],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /adapter-sheetjs\/src\/cells\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/adapter-sheetjs/src'),
          }));
        },
      },
    ],
  });
  hashes[name] = createHash('sha256').update(result.outputFiles[0].contents).digest('hex');
  variants[name] = await import(
    `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].contents).toString('base64')}`
  );
}
const expressions = ['2^3^2', '-2^2'];
const expected = { baseline: expressions, current: ['((2^3)^2)', '((-(2))^2)'] };
const expression = (row, column) =>
  distinct ? `${expressions[column]}+${row}` : expressions[column];
const expectedFormula = (variant, row, column) =>
  distinct
    ? variant === 'baseline'
      ? expression(row, column)
      : `(${expected.current[column]}+${row})`
    : expected[variant][column];
const rows = 10_000;
const input = {
  SheetNames: ['Data'],
  Sheets: {
    Data: {
      '!data': Array.from({ length: rows }, (_, row) =>
        expressions.map((_, column) => ({ t: 'n', f: expression(row, column), v: 0 })),
      ),
      '!ref': `A1:B${rows}`,
    },
  },
};
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  distinct,
  scope: `In-process Node adapter conversion of 20,000 ${distinct ? 'distinct' : 'repeated'} short power formulas, six alternating rounds. The reconstructed baseline skips rewriting and therefore does not preserve spreadsheet precedence; both variants verify every stored formula outside timing. Excludes XLSX parsing, browser Worker, UI, peak memory, other formulas and hardware.`,
  samples: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (let round = 0; round < 6; round++) {
  for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
    const start = performance.now();
    const result = variants[variant].fromSheetJS(input);
    const ms = performance.now() - start;
    let count = 0;
    for (const page of result.file.sheets[0].blocks) {
      const formulas = page.data.formulas;
      if (!formulas) continue;
      for (let i = 0; i < formulas.values.length; i++) {
        const offset = formulas.positions[i],
          row = page.row * 64 + Math.floor(offset / 32);
        const column = page.column * 32 + (offset % 32);
        if (
          row >= rows ||
          column >= 2 ||
          formulas.values[i] !== expectedFormula(variant, row, column)
        )
          throw new Error(`Incorrect ${variant} formula at ${row},${column}`);
        count++;
      }
    }
    if (count !== rows * 2) throw new Error(`Expected ${rows * 2} formulas, got ${count}`);
    report.samples.push({ round: round + 1, variant, ms, verified: count });
    save();
    console.log(JSON.stringify(report.samples.at(-1)));
  }
}
report.finishedAt = new Date().toISOString();
save();
