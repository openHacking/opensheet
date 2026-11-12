import { build } from 'esbuild';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { cpus, release, tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

if (process.argv[2] === '--child') {
  const { fromSheetJS, FileReader } = await import(pathToFileURL(process.argv[3]).href);
  const rows = 20_000;
  const columns = 10;
  const mode = process.argv[4];
  if (!['dense', 'sparse'].includes(mode)) throw new Error('Invalid input representation');
  let data = Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, col) => ({ t: 'n', v: row * columns + col })),
  );
  const worksheet =
    mode === 'dense' ? { '!data': data, '!ref': 'A1:J20000' } : { '!ref': 'A1:J20000' };
  if (mode === 'sparse')
    for (let row = 0; row < rows; row++)
      for (let col = 0; col < columns; col++)
        worksheet[`${String.fromCharCode(65 + col)}${row + 1}`] = data[row][col];
  data = undefined;
  const input = {
    SheetNames: ['Data'],
    Sheets: { Data: worksheet },
  };
  global.gc();
  const before = process.memoryUsage();
  const start = performance.now();
  const { file } = fromSheetJS(input);
  const elapsedMs = performance.now() - start;
  const after = process.memoryUsage();
  const maxRssKiB = process.resourceUsage().maxRSS;
  const reader = new FileReader(file);
  const id = file.sheets[0].id;
  for (let row = 0; row < rows; row++)
    for (let col = 0; col < columns; col++) {
      const value = reader.getCell(id, row, col)?.input;
      if (value?.type !== 'number' || value.value !== row * columns + col)
        throw new Error(`Incorrect cell ${row},${col}`);
    }
  reader.dispose();
  process.stdout.write(
    JSON.stringify({ elapsedMs, before, after, maxRssKiB, verified: rows * columns }),
  );
  process.exit(0);
}

const output = process.argv[2];
if (!output)
  throw new Error('Usage: node scripts/sheetjs-conversion-memory-benchmark.mjs NEW_REPORT');
const importPath = resolve('packages/adapter-sheetjs/src/import.ts');
const source = readFileSync(importPath, 'utf8');
const begin = source.indexOf('    const visitCells = ');
const finish = source.indexOf('    visitCells(collect);', begin);
if (begin < 0 || finish < 0) throw new Error('Cannot reconstruct previous traversal');
const previousTraversal = `    if (hasDense) {
      if (!Array.isArray(dense)) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid dense worksheet data');
      if (dense.length > LIMITS.rows) throw new OpenSheetError('LIMIT_EXCEEDED', 'Dense worksheet too large');
      for (let row = 0; row < dense.length; row++) {
        if (dense[row] === undefined || dense[row] === null) continue;
        if (!Array.isArray(dense[row])) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid dense worksheet row');
        if (dense[row].length > LIMITS.columns) throw new OpenSheetError('LIMIT_EXCEEDED', 'Too many columns');
        for (let col = 0; col < dense[row].length; col++) collect(row, col, dense[row][col]);
      }
    } else for (const [a, c] of Object.entries(ws)) {
      if (a.startsWith('!')) continue;
      const p = parseAddress(a);
      collect(p.row, p.column, c);
    }
`;
let previous =
  source.slice(0, begin) +
  previousTraversal +
  source.slice(finish + '    visitCells(collect);\n'.length);
function replaceOnce(text, from, to) {
  if (text.split(from).length !== 2)
    throw new Error(`Cannot reconstruct previous implementation: ${from}`);
  return text.replace(from, to);
}
previous = replaceOnce(
  previous,
  '    let maxRow = 1,',
  '    const entries: Array<[number, number, any]> = [];\n    let maxRow = 1,',
);
previous = replaceOnce(
  previous,
  '      maxRow = Math.max(maxRow, row + 1);',
  '      entries.push([row, col, c]);\n      maxRow = Math.max(maxRow, row + 1);',
);
previous = replaceOnce(
  previous,
  '    visitCells((row, col, c) => {\n      if (c === undefined || c === null) return;',
  '    for (const [row, col, c] of entries) {',
);
previous = replaceOnce(previous, '    });\n    sh.blocks =', '    }\n    sh.blocks =');
const temp = mkdtempSync(join(tmpdir(), 'opensheet-memory-'));
try {
  for (const [variant, code] of [
    ['previous', previous],
    ['current', source],
  ]) {
    await build({
      stdin: {
        contents: `export { fromSheetJS } from ${JSON.stringify(resolve('packages/adapter-sheetjs/src/index.ts'))}; export { FileReader } from ${JSON.stringify(resolve('packages/core/src/index.ts'))};`,
        resolveDir: process.cwd(),
        loader: 'ts',
      },
      plugins: [
        {
          name: 'variant',
          setup: (builder) =>
            builder.onLoad({ filter: /adapter-sheetjs\/src\/import\.ts$/ }, () => ({
              contents: code,
              loader: 'ts',
              resolveDir: resolve('packages/adapter-sheetjs/src'),
            })),
        },
      ],
      bundle: true,
      platform: 'node',
      format: 'esm',
      outfile: join(temp, `${variant}.mjs`),
    });
  }
  const report = {
    startedAt: new Date().toISOString(),
    source: sourceFingerprint(),
    cpu: cpus()[0]?.model,
    platform: `${process.platform}/${process.arch}/${release()}`,
    node: process.version,
    limitations:
      'Separate Node processes, reconstructed previous traversal, dense and address-keyed synthetic numeric input. Peak RSS includes module loading and fixture allocation; no XLSX parse, browser, Worker or isolated adapter heap peak.',
    samples: [],
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  for (const mode of ['dense', 'sparse'])
    for (let round = 0; round < 3; round++) {
      for (const variant of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
        const child = spawnSync(
          process.execPath,
          [
            '--expose-gc',
            fileURLToPath(import.meta.url),
            '--child',
            join(temp, `${variant}.mjs`),
            mode,
          ],
          { encoding: 'utf8', maxBuffer: 1024 * 1024 },
        );
        if (child.status !== 0) throw new Error(child.stderr || `Child failed: ${child.status}`);
        report.samples.push({ mode, round: round + 1, variant, ...JSON.parse(child.stdout) });
        writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
        console.log(JSON.stringify(report.samples.at(-1)));
      }
    }
} finally {
  rmSync(temp, { recursive: true, force: true });
}
