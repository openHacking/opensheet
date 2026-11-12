import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output)
  throw new Error('Usage: node scripts/conditional-format-formula-benchmark.mjs NEW_REPORT');
const bundle = await build({
  stdin: {
    contents: `export { conditionalViewportWithFormulas, validateConditionalRules } from ${JSON.stringify(resolve('packages/core/src/conditional-format.ts'))}; export { AsyncCalculation } from ${JSON.stringify(resolve('packages/core/src/async-calculation.ts'))};`,
    resolveDir: process.cwd(),
    loader: 'ts',
  },
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
  logLevel: 'silent',
});
const { conditionalViewportWithFormulas, validateConditionalRules, AsyncCalculation } =
  await import(
    `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
  );
const rules = validateConditionalRules(
  [
    {
      id: 'positive',
      range: { startRow: 0, endRow: 1000, startColumn: 0, endColumn: 1 },
      condition: { kind: 'formula', expression: 'A1>0' },
      style: { bold: true },
      stopIfTrue: false,
    },
  ],
  1000,
  1,
);
const cells = Array.from({ length: 1000 }, (_, row) => ({ row, column: 0, value: row, base: {} }));
const calc = new AsyncCalculation(
  async (_sheet, row, column) =>
    column === 0 && row < 1000
      ? { rowId: row, columnId: column, input: { type: 'number', value: row } }
      : undefined,
  () => undefined,
);
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'In-process Node prototype: one simple relative-reference predicate over 1,000 in-memory numeric cells, fresh bounded calculation session each round. Includes rule revalidation, input snapshot and asynchronous evaluation; excludes browser, storage, rendering, memory peak, complex formulas and large viewports.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 5; round++) {
  const session = calc.validationSession().forSheet('data');
  const start = performance.now();
  const values = await conditionalViewportWithFormulas(rules, cells, session.evaluateFormula);
  const elapsedMs = performance.now() - start;
  if (values.length !== 1000 || values.some((style, row) => !!style.bold !== row > 0))
    throw new Error('Incorrect formula appearance');
  report.samples.push({ round: round + 1, elapsedMs, formulas: 1000, verified: values.length });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report.samples.at(-1)));
}
