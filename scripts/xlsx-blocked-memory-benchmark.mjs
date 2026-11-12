import { build } from 'esbuild';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { cpus, release, tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

if (process.argv[2] === '--child') {
  const { fromSheetJS, buildSheetJS, FileReader } = await import(
    pathToFileURL(process.argv[3]).href
  );
  const [mode, position] = process.argv.slice(4),
    rows = 10000,
    columns = 10;
  if (!['retain', 'discard'].includes(mode) || !['first', 'last'].includes(position))
    throw new Error('Invalid case');
  let data = Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, col) => ({ t: 'n', v: row * columns + col })),
  );
  const targetRow = position === 'first' ? 0 : rows - 1,
    targetCol = position === 'first' ? 0 : columns - 1;
  data[targetRow][targetCol] = { t: 's', v: 'x'.repeat(32768) };
  const input = { SheetNames: ['Data'], Sheets: { Data: { '!data': data, '!ref': 'A1:J10000' } } };
  const file = fromSheetJS(input).file;
  input.Sheets.Data['!data'] = [];
  data = null;
  global.gc();
  const before = process.memoryUsage(),
    start = performance.now();
  const result = buildSheetJS(file, {}, true, false, mode === 'discard');
  const elapsedMs = performance.now() - start;
  global.gc();
  const after = process.memoryUsage(),
    maxRssKiB = process.resourceUsage().maxRSS;
  const expectedAddress = position === 'first' ? 'A1' : 'J10000';
  if (
    result.xlsxTextTooLong?.address !== expectedAddress ||
    result.report.summary.blocked !== 1 ||
    result.report.summary.exact !== rows * columns
  )
    throw new Error('Incorrect report');
  const sheet = result.workbook.Sheets.Data;
  if (mode === 'retain') {
    for (let row = 0; row < rows; row++)
      for (let col = 0; col < columns; col++) {
        const value = sheet[`${String.fromCharCode(65 + col)}${row + 1}`]?.v;
        if (
          row === targetRow && col === targetCol
            ? value !== 'x'.repeat(32768)
            : value !== row * columns + col
        )
          throw new Error('Incorrect cell');
      }
  } else if (Object.keys(sheet).some((key) => !key.startsWith('!')))
    throw new Error('Cells were retained');
  const reader = new FileReader(file);
  let verified = 0;
  for (const cell of reader.cells(file.sheets[0].id)) {
    const isTarget = cell.rowId === targetRow && cell.columnId === targetCol;
    if (
      isTarget
        ? cell.input.type !== 'string' || cell.input.value !== 'x'.repeat(32768)
        : cell.input.type !== 'number' || cell.input.value !== cell.rowId * columns + cell.columnId
    )
      throw new Error('Input cell changed');
    verified++;
  }
  reader.dispose();
  if (verified !== rows * columns) throw new Error('Missing input cells');
  process.stdout.write(
    JSON.stringify({
      mode,
      position,
      elapsedMs,
      before,
      after,
      maxRssKiB,
      verified,
      report: {
        summary: result.report.summary,
        issues: result.report.issues.map(({ sheetId, ...issue }) => issue),
        omitted: result.report.omittedIssueOccurrences,
      },
    }),
  );
  process.exit(0);
}

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/xlsx-blocked-memory-benchmark.mjs NEW_REPORT');
const temp = mkdtempSync(join(tmpdir(), 'opensheet-blocked-memory-'));
try {
  const bundle = join(temp, 'adapter.mjs');
  await build({
    stdin: {
      contents: `export { fromSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/import.ts'))}; export { buildSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/export.ts'))}; export { FileReader } from ${JSON.stringify(resolve('packages/core/src/snapshot-reader.ts'))};`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    bundle: true,
    platform: 'node',
    format: 'esm',
    outfile: bundle,
    logLevel: 'silent',
  });
  const report = {
    startedAt: new Date().toISOString(),
    source: sourceFingerprint(),
    cpu: cpus()[0]?.model,
    platform: `${process.platform}/${process.arch}/${release()}`,
    node: process.version,
    scope:
      'Separate Node processes; 100k synthetic cells, one overlong text first or last. Retain and discard are modes of the same build, not a before/after release comparison. Heap is sampled after explicit GC; peak RSS includes fixture construction. Excludes workbook snapshot, XLSX writer, browser Worker and other workloads.',
    samples: [],
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  for (let round = 0; round < 3; round++)
    for (const position of ['first', 'last'])
      for (const mode of round % 2 ? ['discard', 'retain'] : ['retain', 'discard']) {
        const child = spawnSync(
          process.execPath,
          ['--expose-gc', fileURLToPath(import.meta.url), '--child', bundle, mode, position],
          { encoding: 'utf8', maxBuffer: 1024 * 1024 },
        );
        if (child.status !== 0) throw new Error(child.stderr || `Child failed: ${child.status}`);
        report.samples.push({ round: round + 1, ...JSON.parse(child.stdout) });
        writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
        console.log(JSON.stringify(report.samples.at(-1)));
      }
  for (const position of ['first', 'last']) {
    const items = report.samples.filter((sample) => sample.position === position);
    if (items.some((sample) => JSON.stringify(sample.report) !== JSON.stringify(items[0].report)))
      throw new Error('Report mismatch');
  }
} finally {
  rmSync(temp, { recursive: true, force: true });
}
