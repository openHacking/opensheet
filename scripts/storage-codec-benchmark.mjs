import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { performance } from 'node:perf_hooks';
const bundle = await build({
  stdin: {
    contents: `export {createWorkbookFile,encodeFileCells,decodeFileCells} from './packages/core/src/file-codec.ts';export {encodeBlock,decodeBlock} from './packages/core/src/block-codec.ts';export {fixtureValue} from './packages/core/src/engine.ts';export {writeBinary} from './packages/core/src/binary-file.ts';`,
    resolveDir: process.cwd(),
  },
  bundle: true,
  format: 'esm',
  platform: 'node',
  write: false,
});
const {
  createWorkbookFile,
  encodeFileCells,
  decodeFileCells,
  encodeBlock,
  decodeBlock,
  fixtureValue,
  writeBinary,
} = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].text).toString('base64')}`
);
const utf8 = (value) => new TextEncoder().encode(JSON.stringify(value)).length;
const results = {
  date: new Date().toISOString(),
  runtime: process.version,
  fixture:
    '100 columns; fixed seed; 60% Float64 numbers, 30% 32-byte ASCII strings, 10% booleans; no history; one page generated/validated at a time',
  measurements: {},
};
for (const count of [100000, 1000000]) {
  const rows = Math.ceil(count / 100),
    metadata = createWorkbookFile({ sheets: [{ name: 'Storage probe', rows, columns: 100 }] }),
    sheet = metadata.sheets[0];
  let pages = 0,
    rawPageBytes = 0,
    jsonBytes = utf8(metadata),
    encodeMs = 0,
    decodeValidateMs = 0;
  async function* generate(measure = false) {
    for (let r = 0; r < rows; r += 64)
      for (let c = 0; c < 100; c += 32) {
        const cells = {};
        for (let row = r; row < Math.min(r + 64, rows); row++)
          for (let col = c; col < Math.min(c + 32, 100); col++) {
            const n = row * 100 + col,
              v = fixtureValue(n),
              input =
                typeof v === 'number'
                  ? { type: 'number', value: v }
                  : typeof v === 'boolean'
                    ? { type: 'boolean', value: v }
                    : { type: 'string', value: v };
            cells[`${row}:${col}`] = { rowId: row, columnId: col, input };
          }
        let start = performance.now();
        const bytes = encodeBlock(cells, [], [], {});
        const data = encodeFileCells(Object.values(cells), new Map());
        const block = { row: Math.floor(r / 64), column: Math.floor(c / 32), data };
        if (measure) {
          encodeMs += performance.now() - start;
          rawPageBytes += bytes.length;
          jsonBytes += utf8(block) + (pages ? 1 : 0);
          pages++;
          start = performance.now();
          const decoded = decodeBlock(bytes);
          decodeFileCells(block, []);
          for (const cell of Object.values(decoded.cells)) {
            const expected = fixtureValue(cell.rowId * 100 + cell.columnId);
            if (cell.input.value !== expected) throw new Error('Fixture verification failed');
          }
          decodeValidateMs += performance.now() - start;
        }
        yield { sheet: 0, row: block.row, column: block.column, bytes };
      }
  }
  const start = performance.now();
  let binaryFileBytes = 0;
  for await (const bytes of writeBinary(metadata, generate(true))) binaryFileBytes += bytes.length;
  results.measurements[count] = {
    pages,
    jsonBytes,
    rawPageBytes,
    binaryFileBytes,
    encodeMs: +encodeMs.toFixed(2),
    decodeValidateMs: +decodeValidateMs.toFixed(2),
    binaryEncodeCompressionAndVerificationMs: +(performance.now() - start).toFixed(2),
  };
}
results.emptyMaxDimensionsJSONBytes = utf8(
  createWorkbookFile({ sheets: [{ name: 'Empty', rows: 100000, columns: 1000 }] }),
);
writeFileSync('benchmarks/storage-binary.json', JSON.stringify(results, null, 2) + '\n');
console.log(JSON.stringify(results, null, 2));
