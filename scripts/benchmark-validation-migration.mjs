import { build } from 'esbuild';
import { availableParallelism, cpus, platform, release } from 'node:os';

const rows = Number(process.argv[2] ?? 31_250);
const operation = process.argv[3] ?? 'upgrade';
const columns = Number(process.argv[4] ?? 32);
if (
  !Number.isInteger(rows) ||
  rows < 1 ||
  !Number.isInteger(columns) ||
  columns < 32 ||
  columns % 32 !== 0 ||
  !['upgrade', 'downgrade', 'stage'].includes(operation)
) {
  throw new Error(
    'Usage: node --expose-gc scripts/benchmark-validation-migration.mjs [rows] [upgrade|downgrade|stage] [columns in multiples of 32]',
  );
}
if (!globalThis.gc) throw new Error('Run Node with --expose-gc');

const bundle = await build({
  stdin: {
    contents: `export { createWorkbookFile } from './packages/core/src/file-codec.ts';
      export { upgradeValidationFile, downgradeValidationFile, stageValidationUpgrade } from './packages/core/src/validation-file-migration.ts';`,
    resolveDir: process.cwd(),
  },
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
});
const {
  createWorkbookFile,
  upgradeValidationFile,
  downgradeValidationFile,
  stageValidationUpgrade,
} = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
);
const source = createWorkbookFile({ sheets: [{ name: 'Dense', rows, columns }] });
for (let row = 0; row < rows; row += 64) {
  const count = Math.min(64, rows - row) * 32;
  for (let column = 0; column < columns; column += 32)
    source.sheets[0].blocks.push({
      row: row / 64,
      column: column / 32,
      data: {
        numbers: {
          positions: Array.from({ length: count }, (_, index) => index),
          values: Array.from(
            { length: count },
            (_, index) => (row + Math.floor(index / 32)) * columns + column + (index % 32),
          ),
        },
      },
    });
}
const input =
  operation !== 'downgrade'
    ? source
    : {
        ...source,
        schemaVersion: 4,
        sheets: source.sheets.map((sheet) => ({ ...sheet, validation: [] })),
      };
globalThis.gc();
const rssBeforeBytes = process.memoryUsage().rss;
const peakBeforeKiB = process.resourceUsage().maxRSS;
const start = performance.now();
let staged = 0;
let lastValue;
const result =
  operation === 'stage'
    ? await stageValidationUpgrade(input, async (_sheet, block) => {
        staged++;
        lastValue = block.data.numbers.values.at(-1);
      })
    : operation === 'upgrade'
      ? upgradeValidationFile(input)
      : downgradeValidationFile(input);
const elapsedMs = performance.now() - start;
const peakAfterKiB = process.resourceUsage().maxRSS;
if (
  result.schemaVersion !== (operation === 'downgrade' ? 3 : 4) ||
  (operation === 'stage' ? staged : result.sheets[0].blocks.length) !==
    (Math.ceil(rows / 64) * columns) / 32 ||
  (operation === 'stage'
    ? lastValue
    : result.sheets[0].blocks.at(-1).data.numbers.values.at(-1)) !==
    rows * columns - 1
) {
  throw new Error('Migration lost fixture data');
}
console.log(
  JSON.stringify({
    operation,
    cells: rows * columns,
    rows,
    columns,
    elapsedMs,
    rssBeforeBytes,
    peakBeforeKiB,
    peakAfterKiB,
    node: process.version,
    platform: platform(),
    release: release(),
    arch: process.arch,
    cpu: cpus()[0]?.model,
    availableParallelism: availableParallelism(),
  }),
);
