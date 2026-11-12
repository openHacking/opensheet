import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/defined-name-scan-benchmark.mjs NEW_REPORT');
const source = readFileSync('packages/adapter-sheetjs/src/import.ts', 'utf8');
const scan = `      if (
        record.input.type === 'formula' &&
        referencesDefinedName(
          record.input.expression,
          nameIndex.global,
          nameIndex.local.get(index),
          qualifiedLocalNames,
        )
      )
        issue(
          'DEFINED_NAME_REFERENCE_DROPPED',
          'defined-name-reference',
          'Formula references a defined name that is not preserved',
          'dropped',
          address(row, col),
        );
`;
if (source.split(scan).length !== 2) throw new Error('Cannot reconstruct control');
async function compile(code) {
  const result = await build({
    entryPoints: [resolve('packages/adapter-sheetjs/src/index.ts')],
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /adapter-sheetjs\/src\/import\.ts$/ }, () => ({
            contents: code,
            loader: 'ts',
            resolveDir: resolve('packages/adapter-sheetjs/src'),
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
    `data:text/javascript;base64,${Buffer.from(result.outputFiles[0].contents).toString('base64')}`
  );
}
const variants = {
  control: await compile(source.replace(scan, '')),
  current: await compile(source),
};
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  cells: 30000,
  scope:
    'Alternating in-process Node adapter conversions with synthetic repeated formulas. Control removes only the new per-formula name scan; name-index construction remains. No XLSX parsing, browser Worker, peak memory or end-to-end UI.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (const scenario of ['no-names', 'unmatched', 'matched', 'qualified', 'qualified-global']) {
  const formulas =
    scenario === 'qualified' || scenario === 'qualified-global'
      ? ['Data!Budget+1', 'SUM(Data!Budget,1)', 'Data!A1']
      : scenario === 'matched'
        ? ['Budget+1', 'IF(TRUE,Budget,0)', 'SUM(A1:A2)']
        : ['SUM(A1:A2)', 'IF(TRUE,ROUND(1.25,1),0)', 'CONCAT("Budget", "x")'];
  const input = {
    SheetNames: ['Data'],
    Sheets: {
      Data: {
        '!data': Array.from({ length: 10000 }, () => formulas.map((f) => ({ t: 'n', f, v: 0 }))),
        '!ref': 'A1:C10000',
      },
    },
    Workbook: {
      Names:
        scenario === 'no-names'
          ? []
          : [
              {
                Name: 'Budget',
                Ref: 'Data!$A$1',
                ...(scenario === 'qualified' ? { Sheet: 0 } : {}),
              },
            ],
    },
  };
  for (let round = 0; round < 4; round++)
    for (const variant of round % 2 ? ['current', 'control'] : ['control', 'current']) {
      const start = performance.now(),
        result = variants[variant].fromSheetJS(input),
        elapsedMs = performance.now() - start;
      let verified = 0;
      for (const page of result.file.sheets[0].blocks) {
        const cells = page.data.formulas;
        if (!cells) continue;
        for (let i = 0; i < cells.values.length; i++) {
          const col = page.column * 32 + (cells.positions[i] % 32);
          if (col >= formulas.length || cells.values[i] !== formulas[col])
            throw new Error('Formula mismatch');
          verified++;
        }
      }
      const issue = result.report.issues.find(
        (item) => item.code === 'DEFINED_NAME_REFERENCE_DROPPED',
      );
      if (
        verified !== report.cells ||
        (variant === 'current' &&
        (scenario === 'matched' || scenario === 'qualified' || scenario === 'qualified-global')
          ? issue?.count !== 20000
          : issue !== undefined)
      )
        throw new Error('Incorrect report or cell count');
      report.samples.push({ scenario, round: round + 1, variant, elapsedMs, verified });
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    }
}
console.log(JSON.stringify(report.samples));
