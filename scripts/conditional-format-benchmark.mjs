import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/conditional-format-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: [resolve('packages/core/src/conditional-format.ts')],
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
  logLevel: 'silent',
});
const { conditionalViewport, validateConditionalRules } = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
);
const rules = validateConditionalRules(
  Array.from({ length: 1000 }, (_, i) => ({
    id: `r${i}`,
    range: { startRow: 0, endRow: 1000, startColumn: 0, endColumn: 1 },
    condition: { kind: 'number', operator: 'gte', first: 0 },
    style: { color: i === 0 ? '#ff0000' : '#0000ff' },
    stopIfTrue: false,
  })),
  1000,
  1,
);
const cells = Array.from({ length: 1000 }, (_, row) => ({
  row,
  column: 0,
  value: row,
  base: { background: '#ffffff' },
}));
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'In-process Node prototype only; 1,000 loaded numeric cells and 1,000 fully overlapping matching rules, no stop conditions. Validation and fixture construction are excluded. Does not measure browser rendering, storage, memory peak, text or formula rules.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 5; round++) {
  const start = performance.now();
  const values = conditionalViewport(rules, cells);
  const elapsedMs = performance.now() - start;
  if (
    values.length !== 1000 ||
    values.some((style) => style.color !== '#ff0000' || style.background !== '#ffffff')
  )
    throw new Error('Incorrect conditional formatting result');
  report.samples.push({
    round: round + 1,
    elapsedMs,
    comparisons: 1_000_000,
    verified: values.length,
  });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report.samples.at(-1)));
}
