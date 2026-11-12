import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/unicode-name-scan-benchmark.mjs NEW_REPORT');
const current = readFileSync('packages/adapter-sheetjs/src/defined-name-scan.ts', 'utf8');
const before = current
  .replace('/^[\\p{L}_\\\\][\\p{L}\\p{N}\\p{M}_.]*$/u', '/^[A-Za-z_\\\\][A-Za-z_0-9.]*$/')
  .replaceAll('/[\\p{L}_\\\\][\\p{L}\\p{N}\\p{M}_.]*/uy', '/[A-Za-z_\\\\][A-Za-z_0-9.]*/y');
if (before === current || before.includes('\\p{L}'))
  throw new Error('Cannot reconstruct ASCII control');
async function compile(source) {
  const result = await build({
    entryPoints: [resolve('packages/adapter-sheetjs/src/index.ts')],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
    plugins: [
      {
        name: 'scanner',
        setup(builder) {
          builder.onLoad({ filter: /adapter-sheetjs\/src\/defined-name-scan\.ts$/ }, () => ({
            contents: source,
            loader: 'ts',
          }));
        },
      },
    ],
  });
  return import(
    `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].contents).toString('base64')}`
  );
}
const variants = { before: await compile(before), current: await compile(current) };
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  cells: 30000,
  scope:
    'Alternating local Node conversions of identical ASCII formulas. The control uses ASCII-only identifier regular expressions while retaining the backslash prefix and boundary checks. No XLSX parsing, browser Worker, UI, peak memory or Unicode throughput.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (const scenario of ['no-names', 'qualified']) {
  const formulas =
    scenario === 'qualified'
      ? ['Data!Budget+1', 'SUM(Data!Budget,1)', 'Data!A1']
      : ['SUM(A1:A2)', 'IF(TRUE,ROUND(1.25,1),0)', 'CONCAT("Budget","x")'];
  const input = {
    SheetNames: ['Data'],
    Sheets: {
      Data: {
        '!data': Array.from({ length: 10000 }, () => formulas.map((f) => ({ t: 'n', f, v: 0 }))),
        '!ref': 'A1:C10000',
      },
    },
    Workbook: { Names: scenario === 'qualified' ? [{ Name: 'Budget', Ref: 'Data!$A$1' }] : [] },
  };
  for (let round = 0; round < 4; round++)
    for (const variant of round % 2 ? ['current', 'before'] : ['before', 'current']) {
      const start = performance.now(),
        result = variants[variant].fromSheetJS(input),
        elapsedMs = performance.now() - start;
      let verified = 0;
      for (const page of result.file.sheets[0].blocks)
        if (page.data.formulas)
          for (let i = 0; i < page.data.formulas.values.length; i++) {
            const col = page.column * 32 + (page.data.formulas.positions[i] % 32);
            if (page.data.formulas.values[i] !== formulas[col]) throw new Error('Formula mismatch');
            verified++;
          }
      const issue = result.report.issues.find(
        (item) => item.code === 'DEFINED_NAME_REFERENCE_DROPPED',
      );
      if (
        verified !== 30000 ||
        (scenario === 'qualified' ? issue?.count !== 20000 : issue !== undefined)
      )
        throw new Error('Report mismatch');
      report.samples.push({ scenario, round: round + 1, variant, elapsedMs, verified });
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    }
}
console.log(JSON.stringify(report.samples));
