import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/conditional-stage-benchmark.mjs NEW_REPORT');
const bundle = await build({
  stdin: {
    contents: `export { fromSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/import.ts'))}; export { upgradeValidationFile, upgradeConditionalFile, stageConditionalFile } from ${JSON.stringify(resolve('packages/core/src/validation-file-migration.ts'))};`,
    resolveDir: process.cwd(),
    loader: 'ts',
  },
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
  logLevel: 'silent',
});
const { fromSheetJS, upgradeValidationFile, upgradeConditionalFile, stageConditionalFile } =
  await import(
    `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
  );
const rows = 20000,
  columns = 10;
const data = Array.from({ length: rows }, (_, row) =>
  Array.from({ length: columns }, (_, column) => ({ t: 'n', v: row * columns + column })),
);
const v4 = upgradeValidationFile(
  fromSheetJS({ SheetNames: ['Data'], Sheets: { Data: { '!data': data, '!ref': 'A1:J20000' } } })
    .file,
);
v4.sheets[0].validation.push({
  id: 'positive',
  range: { startRow: 0, endRow: rows, startColumn: 0, endColumn: columns },
  rule: { kind: 'number', operator: 'gte', first: 0, allowBlank: false, policy: 'reject' },
});
const file = upgradeConditionalFile(v4);
file.sheets[0].conditionalFormats.push({
  id: 'highlight',
  range: { startRow: 0, endRow: rows, startColumn: 0, endColumn: columns },
  condition: { kind: 'formula', expression: 'A1>0' },
  style: { bold: true },
  stopIfTrue: true,
});
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  scope:
    'In-process Node prototype: stages 200,000 synthetic numeric cells and two rule entries to an in-memory sink. Timing excludes fixture generation and full page verification. No IndexedDB, Worker, crash recovery, browser responsiveness or peak memory measurement.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 5; round++) {
  const pages = [],
    start = performance.now();
  const metadata = await stageConditionalFile(file, async (sheet, block) =>
    pages.push({ sheet, block }),
  );
  const elapsedMs = performance.now() - start;
  if (
    metadata.schemaVersion !== 5 ||
    metadata.sheets[0].validation.length !== 1 ||
    metadata.sheets[0].conditionalFormats.length !== 1
  )
    throw new Error('Incorrect metadata');
  let verified = 0;
  for (const { sheet, block } of pages) {
    if (sheet !== 0) throw new Error('Unexpected sheet');
    const numeric = block.data.numbers;
    if (!numeric) continue;
    for (let index = 0; index < numeric.positions.length; index++) {
      const offset = numeric.positions[index],
        row = block.row * 64 + Math.floor(offset / 32),
        column = block.column * 32 + (offset % 32);
      if (numeric.values[index] !== row * columns + column)
        throw new Error(`Incorrect cell ${row},${column}`);
      verified++;
    }
  }
  if (verified !== rows * columns)
    throw new Error(`Expected ${rows * columns} cells, got ${verified}`);
  report.samples.push({ round: round + 1, elapsedMs, pages: pages.length, verified });
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  console.log(JSON.stringify(report.samples.at(-1)));
}
