import { build } from 'esbuild';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { cpus, release, tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

if (process.argv[2] === '--child') {
  const { fromSheetJS, buildSheetJS } = await import(pathToFileURL(process.argv[3]).href);
  const mode = process.argv[4],
    rows = 20000,
    columns = 10;
  if (!['full', 'report-only'].includes(mode)) throw new Error('Invalid mode');
  let data = Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, col) => ({ t: 'n', v: row * columns + col })),
  );
  const input = { SheetNames: ['Data'], Sheets: { Data: { '!data': data, '!ref': 'A1:J20000' } } };
  const file = fromSheetJS(input).file;
  input.Sheets.Data['!data'] = [];
  data = null;
  file.sheets[0].freeze = { rows: 1, columns: 1 };
  const prepared = buildSheetJS(file, {}, true);
  global.gc();
  const before = process.memoryUsage(),
    start = performance.now();
  const result = buildSheetJS(file, {}, false, mode === 'report-only');
  const elapsedMs = performance.now() - start,
    after = process.memoryUsage(),
    maxRssKiB = process.resourceUsage().maxRSS;
  if (
    prepared.workbook.Sheets.Data.A1.v !== 0 ||
    result.report.summary.exact !== rows * columns ||
    result.report.summary.dropped !== 1 ||
    result.report.issues[0]?.code !== 'VIEW_NOT_EXPORTED'
  )
    throw new Error('Invalid retained workbook or report');
  const sheet = result.workbook.Sheets.Data;
  if (mode === 'full') {
    for (let row = 0; row < rows; row++)
      for (let col = 0; col < columns; col++)
        if (sheet[`${String.fromCharCode(65 + col)}${row + 1}`]?.v !== row * columns + col)
          throw new Error('Incorrect cell');
  } else if (Object.keys(sheet).some((key) => !key.startsWith('!')))
    throw new Error('Unexpected retained cell');
  process.stdout.write(
    JSON.stringify({ elapsedMs, before, after, maxRssKiB, verified: rows * columns }),
  );
  process.exit(0);
}

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/xlsx-fallback-memory-benchmark.mjs NEW_REPORT');
const temp = mkdtempSync(join(tmpdir(), 'opensheet-fallback-memory-'));
try {
  const bundle = join(temp, 'adapter.mjs');
  await build({
    stdin: {
      contents: `export { fromSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/import.ts'))}; export { buildSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/export.ts'))};`,
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
      'Separate Node processes with 200k synthetic numeric cells and the first SheetJS export retained. Measures the second report-only or full conversion, excluding XLSX ZIP serialization; peak RSS includes module loading and fixture generation. No browser, Worker, peak isolated heap or other workloads.',
    samples: [],
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  for (let round = 0; round < 3; round++) {
    for (const mode of round % 2 ? ['report-only', 'full'] : ['full', 'report-only']) {
      const child = spawnSync(
        process.execPath,
        ['--expose-gc', fileURLToPath(import.meta.url), '--child', bundle, mode],
        { encoding: 'utf8', maxBuffer: 1024 * 1024 },
      );
      if (child.status !== 0) throw new Error(child.stderr || `Child failed: ${child.status}`);
      report.samples.push({ round: round + 1, mode, ...JSON.parse(child.stdout) });
      writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
      console.log(JSON.stringify(report.samples.at(-1)));
    }
  }
} finally {
  rmSync(temp, { recursive: true, force: true });
}
