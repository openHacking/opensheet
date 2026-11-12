import { createHash } from 'node:crypto';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { cpus, totalmem } from 'node:os';
import { transform } from 'esbuild';
import * as XLSX from 'xlsx';
import { fromSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { FileReader } from '../packages/core/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!output || existsSync(output)) throw new Error('Pass a new --output=report.json path');
const source = readFileSync('apps/playground/src/zip-guard.ts', 'utf8');
const compiled = await transform(source, { loader: 'ts', format: 'esm', target: 'es2022' });
const { validateZip } = await import(
  `data:text/javascript;base64,${Buffer.from(compiled.code).toString('base64')}`
);
const sha256 = (value) => createHash('sha256').update(value).digest('hex');
const report = {
  date: new Date().toISOString(),
  scope:
    'Node import stages on synthetic four-column XLSX, three sequential rounds per size. Not browser/Worker UI latency, a hard memory ceiling, or a universal capacity gate. Fixture generation is outside timings.',
  host: {
    node: process.version,
    platform: process.platform,
    arch: process.arch,
    cpu: cpus()[0].model,
    logicalCpus: cpus().length,
    memoryBytes: totalmem(),
  },
  source: sourceFingerprint(),
  adapterBuild: localBuildFingerprint('packages/adapter-sheetjs/dist'),
  guardSha256: sha256(source),
  sheetjs: XLSX.version,
  cases: [],
};
for (const rows of [1000, 10000, 50000]) {
  const cells = Array.from({ length: rows }, (_, index) => [
    index,
    `record-${index}`,
    index % 7 === 0,
    index / 10,
  ]);
  const workbook = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet(cells), 'Data');
  const bytes = XLSX.write(workbook, { type: 'array', bookType: 'xlsx', compression: true });
  const result = {
    rows,
    columns: 4,
    inputBytes: bytes.byteLength,
    sha256: sha256(new Uint8Array(bytes)),
    rounds: [],
  };
  for (let round = 0; round < 3; round++) {
    const start = performance.now();
    await validateZip(bytes, true);
    const verified = performance.now();
    const parsed = XLSX.read(bytes, {
      type: 'array',
      dense: true,
      cellNF: true,
      cellStyles: true,
      cellFormula: true,
      bookVBA: true,
    });
    const decoded = performance.now();
    const imported = fromSheetJS(parsed);
    const finished = performance.now();
    const data = parsed.Sheets.Data['!data'];
    const reader = new FileReader(imported.file);
    const sheetId = imported.file.sheets[0].id;
    const rowCount = reader.sheetData(sheetId).rowOrder.length;
    if (
      data.length !== rows ||
      data[rows - 1][0].v !== rows - 1 ||
      rowCount !== Math.max(100, rows)
    )
      throw new Error('Import correctness assertion failed');
    for (let row = 0; row < rows; row++) {
      for (let column = 0; column < 4; column++) {
        const actual = reader.getCell(sheetId, row, column)?.input;
        const expected = cells[row][column];
        const expectedType = typeof expected === 'string' ? 'string' : typeof expected;
        if (actual?.type !== expectedType || actual.value !== expected)
          throw new Error(`Incorrect imported cell at ${row},${column}`);
      }
    }
    reader.dispose();
    result.rounds.push({
      guardMs: verified - start,
      parseMs: decoded - verified,
      adapterMs: finished - decoded,
      totalMs: finished - start,
      checksPassed: true,
    });
  }
  report.cases.push(result);
  process.stdout.write(JSON.stringify(result) + '\n');
}
report.processMaxRssKiB = process.resourceUsage().maxRSS;
report.memoryScope =
  'Process high-water mark includes fixture generation, previous rounds, parser, adapter and runtime; not isolated import memory.';
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
