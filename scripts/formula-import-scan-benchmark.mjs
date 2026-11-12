import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/formula-import-scan-benchmark.mjs NEW_REPORT');
const path = resolve('packages/adapter-sheetjs/src/cells.ts');
const source = readFileSync(path, 'utf8');
const scan = `  if (formula) {
    const name = unsupportedFunction(formula);
    if (name)
      issue(
        'UNSUPPORTED_FORMULA_FUNCTION',
        'formula-function',
        'Formula uses a function that OpenSheet does not support',
        'dropped',
        address(row, col),
      );
  }
`;
if (source.split(scan).length !== 2) throw new Error('Cannot reconstruct baseline safely');
async function compile(code) {
  const result = await build({
    entryPoints: [resolve('packages/adapter-sheetjs/src/index.ts')],
    plugins: [
      {
        name: 'variant',
        setup: (builder) =>
          builder.onLoad({ filter: /adapter-sheetjs\/src\/cells\.ts$/ }, () => ({
            contents: code,
            loader: 'ts',
            resolveDir: resolve('packages/adapter-sheetjs/src'),
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
const variants = {
  baseline: await compile(source.replace(scan, '')),
  current: await compile(source),
};
const formulas = ['SUM(1,2)', 'IF(TRUE,ROUND(1.25,1),0)', 'CONCAT("quoted UNKNOWN(1)","text")'];
const data = Array.from({ length: 20_000 }, () => formulas.map((f) => ({ t: 'n', f, v: 0 })));
const input = { SheetNames: ['Data'], Sheets: { Data: { '!data': data, '!ref': 'A1:C20000' } } };
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  cells: 60_000,
  limitations:
    'In-process Node adapter conversion only; alternating variants, synthetic repeated supported formulas, no XLSX parse, browser, Worker or isolated memory measurement.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 6; round++) {
  for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
    const start = performance.now();
    const result = variants[variant].fromSheetJS(input);
    const elapsedMs = performance.now() - start;
    let actual = 0;
    for (const page of result.file.sheets[0].blocks) {
      const column = page.data.formulas;
      if (!column) continue;
      for (let index = 0; index < column.values.length; index++) {
        const offset = column.positions[index];
        const row = page.row * 64 + Math.floor(offset / 32);
        const col = page.column * 32 + (offset % 32);
        if (row >= 20_000 || col >= formulas.length || column.values[index] !== formulas[col])
          throw new Error(`Incorrect formula at ${row},${col}`);
        actual++;
      }
    }
    if (
      actual !== 60_000 ||
      result.report.issues.some((issue) => issue.code === 'UNSUPPORTED_FORMULA_FUNCTION')
    )
      throw new Error('Unexpected conversion result');
    report.samples.push({ round: round + 1, variant, elapsedMs });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report.samples.at(-1)));
  }
}
const longFormula = `CONCAT("${'x'.repeat(32_740)}UNKNOWN(1)")`;
const longInput = {
  SheetNames: ['Data'],
  Sheets: {
    Data: {
      '!data': Array.from({ length: 100 }, () => [{ t: 's', f: longFormula, v: '' }]),
      '!ref': 'A1:A100',
    },
  },
};
report.longSamples = [];
for (let round = 0; round < 3; round++) {
  for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
    const start = performance.now();
    const result = variants[variant].fromSheetJS(longInput);
    const elapsedMs = performance.now() - start;
    const values = result.file.sheets[0].blocks.flatMap((page) => page.data.formulas?.values ?? []);
    if (
      values.length !== 100 ||
      values.some((formula) => formula !== longFormula) ||
      result.report.issues.some((issue) => issue.code === 'UNSUPPORTED_FORMULA_FUNCTION')
    )
      throw new Error('Incorrect long formula conversion');
    report.longSamples.push({ round: round + 1, variant, elapsedMs });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify({ case: 'long-quoted-text', ...report.longSamples.at(-1) }));
  }
}
const distinctInput = {
  SheetNames: ['Data'],
  Sheets: {
    Data: {
      '!data': Array.from({ length: 20_000 }, (_, row) => [
        { t: 'n', f: `UNKNOWN${row}(1)`, v: 0 },
      ]),
      '!ref': 'A1:A20000',
    },
  },
};
report.distinctSamples = [];
for (let round = 0; round < 3; round++) {
  for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
    const start = performance.now();
    const result = variants[variant].fromSheetJS(distinctInput);
    const elapsedMs = performance.now() - start;
    let verified = 0;
    for (const page of result.file.sheets[0].blocks) {
      const formulas = page.data.formulas;
      if (!formulas) continue;
      for (let index = 0; index < formulas.values.length; index++) {
        const offset = formulas.positions[index];
        const row = page.row * 64 + Math.floor(offset / 32);
        if (
          page.column !== 0 ||
          offset % 32 !== 0 ||
          row >= 20_000 ||
          formulas.values[index] !== `UNKNOWN${row}(1)`
        )
          throw new Error(`Incorrect distinct formula at ${row}`);
        verified++;
      }
    }
    const warning = result.report.issues.find(
      (issue) => issue.code === 'UNSUPPORTED_FORMULA_FUNCTION',
    );
    if (
      verified !== 20_000 ||
      (variant === 'current' ? warning?.count !== 20_000 : warning !== undefined)
    )
      throw new Error('Incorrect distinct-function report');
    report.distinctSamples.push({ round: round + 1, variant, elapsedMs, verified });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify({ case: 'distinct-unsupported', ...report.distinctSamples.at(-1) }));
  }
}
const longIdentifier = 'A'.repeat(32_000);
const identifierInput = {
  SheetNames: ['Data'],
  Sheets: {
    Data: {
      '!data': Array.from({ length: 100 }, () => [{ t: 'n', f: longIdentifier, v: 0 }]),
      '!ref': 'A1:A100',
    },
  },
};
report.identifierSamples = [];
for (let round = 0; round < 3; round++) {
  for (const variant of round % 2 ? ['current', 'baseline'] : ['baseline', 'current']) {
    const start = performance.now();
    const result = variants[variant].fromSheetJS(identifierInput);
    const elapsedMs = performance.now() - start;
    const values = result.file.sheets[0].blocks.flatMap((page) => page.data.formulas?.values ?? []);
    if (
      values.length !== 100 ||
      values.some((formula) => formula !== longIdentifier) ||
      result.report.issues.some((issue) => issue.code === 'UNSUPPORTED_FORMULA_FUNCTION')
    )
      throw new Error('Incorrect long identifier conversion');
    report.identifierSamples.push({ round: round + 1, variant, elapsedMs });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(
      JSON.stringify({ case: 'long-non-call-identifier', ...report.identifierSamples.at(-1) }),
    );
  }
}
