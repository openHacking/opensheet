import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/named-range-model-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: [resolve('packages/core/src/named-range-model.ts')],
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
});
const model = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
);
const sheets = [
  { id: 'first', rows: 10, columns: 3 },
  { id: 'second', rows: 10, columns: 3 },
];
const names = [];
for (let i = 0; i < 4096; i++) {
  names.push({
    id: `g_${i}`,
    name: `Range_${i}`,
    scopeSheetId: null,
    target: { sheetId: 'first', range: { startRow: 2, startColumn: 0, endRow: 4, endColumn: 1 } },
  });
  names.push({
    id: `l_${i}`,
    name: `range_${i}`,
    scopeSheetId: 'second',
    target: { sheetId: 'second', range: { startRow: 1, startColumn: 0, endRow: 2, endColumn: 1 } },
  });
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Internal rectangle-only model; one warmed Node process, 8,192 synthetic definitions per round. Index build, 8,192 detached lookups and a row transform are timed separately. Excludes browser, storage, formula evaluation, cancellation and isolated peak memory.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 4; round++) {
  let start = performance.now();
  const index = model.indexNamedRanges(names, sheets),
    indexMs = performance.now() - start;
  start = performance.now();
  for (let i = 0; i < 4096; i++)
    if (
      index.resolve(`RANGE_${i}`, 'first')?.id !== `g_${i}` ||
      index.resolve(`RANGE_${i}`, 'second')?.id !== `l_${i}`
    )
      throw new Error('Incorrect scope');
  const lookupMs = performance.now() - start;
  start = performance.now();
  const shifted = model.transformNamedRanges(names, sheets, 'first', 'row', 'insert', 1, 1),
    transformMs = performance.now() - start;
  for (let i = 0; i < 4096; i++)
    if (
      shifted[i * 2].target.range.startRow !== 3 ||
      shifted[i * 2 + 1].target.range.startRow !== 1 ||
      names[i * 2].target.range.startRow !== 2
    )
      throw new Error('Incorrect transform');
  report.samples.push({ round: round + 1, indexMs, lookupMs, transformMs, verified: 8192 });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
}
console.log(JSON.stringify(report.samples));
