import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { cpus, release, tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

if (process.argv[2] === '--child') {
  const { fromSheetJS, buildSheetJS } = await import(pathToFileURL(process.argv[3]).href);
  const rows = 10_000,
    columns = 10;
  const data = Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, col) => ({ t: 'n', v: row * columns + col })),
  );
  const { file } = fromSheetJS({
    SheetNames: ['Data'],
    Sheets: { Data: { '!data': data, '!ref': 'A1:J10000' } },
  });
  global.gc();
  const start = performance.now();
  const { workbook, report } = buildSheetJS(file, {}, false);
  const elapsedMs = performance.now() - start;
  const sheet = workbook.Sheets.Data;
  for (let row = 0; row < rows; row++)
    for (let col = 0; col < columns; col++) {
      const address = `${String.fromCharCode(65 + col)}${row + 1}`;
      if (sheet[address]?.v !== row * columns + col) throw new Error(`Incorrect ${address}`);
    }
  if (report.summary.exact !== rows * columns) throw new Error('Incorrect report count');
  process.stdout.write(JSON.stringify({ elapsedMs, verified: rows * columns }));
  process.exit(0);
}

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/snapshot-copy-export-benchmark.mjs NEW_REPORT');
const path = resolve('packages/core/src/snapshot-reader.ts');
const current = readFileSync(path, 'utf8');
const original = 'for (const cell of this.rawCells(sheetId)) yield copyCell(cell);';
if (current.split(original).length !== 2)
  throw new Error('Cannot reconstruct prior cells traversal');
const previous = current.replace(
  original,
  'for (const cell of this.rawCells(sheetId)) yield cell;',
);
const temp = mkdtempSync(join(tmpdir(), 'opensheet-copy-export-'));
try {
  const hashes = {};
  for (const [variant, code] of [
    ['previous', previous],
    ['current', current],
  ]) {
    const outfile = join(temp, `${variant}.mjs`);
    await build({
      stdin: {
        contents: `export { fromSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/import.ts'))}; export { buildSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/export.ts'))};`,
        resolveDir: process.cwd(),
        loader: 'ts',
      },
      plugins: [
        {
          name: 'snapshot-variant',
          setup(builder) {
            builder.onLoad({ filter: /core\/src\/snapshot-reader\.ts$/ }, () => ({
              contents: code,
              loader: 'ts',
              resolveDir: resolve('packages/core/src'),
            }));
          },
        },
      ],
      bundle: true,
      platform: 'node',
      format: 'esm',
      outfile,
    });
    hashes[variant] = createHash('sha256').update(readFileSync(outfile)).digest('hex');
  }
  const report = {
    startedAt: new Date().toISOString(),
    source: sourceFingerprint(),
    hashes,
    cpu: cpus()[0]?.model,
    platform: `${process.platform}/${process.arch}/${release()}`,
    node: process.version,
    limitations:
      'Fresh Node processes, one synthetic dense numeric fixture. Previous reconstructs only cells() returning internal objects. Includes reader construction, validation and SheetJS object creation, not XLSX serialization, browsers, memory or UI latency.',
    samples: [],
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  for (let round = 0; round < 4; round++)
    for (const variant of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
      const child = spawnSync(
        process.execPath,
        ['--expose-gc', fileURLToPath(import.meta.url), '--child', join(temp, `${variant}.mjs`)],
        { encoding: 'utf8' },
      );
      if (child.status !== 0) throw new Error(child.stderr || `Child failed: ${child.status}`);
      report.samples.push({ round: round + 1, variant, ...JSON.parse(child.stdout) });
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
      console.log(JSON.stringify(report.samples.at(-1)));
    }
} finally {
  rmSync(temp, { recursive: true, force: true });
}
