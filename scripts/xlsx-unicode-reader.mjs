import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { createWorkbookFile, encodeFileCells } from '../packages/core/dist/index.js';
import { writeXlsx } from '../packages/adapter-sheetjs/dist/index.js';

const prefix = process.argv[2];
if (!prefix) throw new Error('Usage: node scripts/xlsx-unicode-reader.mjs NEW_PREFIX');
const source = resolve(`${prefix}.xlsx`);
const saved = resolve(`${prefix}.saved.xlsx`);
const reportPath = resolve(`${prefix}.json`);
const file = createWorkbookFile({ sheets: [{ name: 'Unicode', rows: 8, columns: 1 }] });
const cells = [
  { rowId: 0, columnId: 0, input: { type: 'string', value: '\uD800' } },
  { rowId: 1, columnId: 0, input: { type: 'string', value: '\uDC00' } },
  { rowId: 2, columnId: 0, input: { type: 'string', value: '🙂' } },
  { rowId: 3, columnId: 0, input: { type: 'string', value: '\uFFFE' } },
  { rowId: 4, columnId: 0, input: { type: 'string', value: 'note' }, note: 'note\uD800' },
  {
    rowId: 5,
    columnId: 0,
    input: { type: 'string', value: 'link' },
    link: { target: 'https://example.com/', tooltip: 'tip\uDC00' },
  },
  {
    rowId: 6,
    columnId: 0,
    input: { type: 'string', value: 'plain link' },
    link: { target: 'https://example.org/', tooltip: 'ordinary tip' },
  },
  {
    rowId: 7,
    columnId: 0,
    input: { type: 'string', value: 'emoji link' },
    link: { target: 'https://example.net/', tooltip: 'tip🙂' },
  },
];
file.sheets[0].blocks = [{ row: 0, column: 0, data: encodeFileCells(cells, new Map()) }];
const generated = writeXlsx(file, XLSX);
const temp = mkdtempSync(join(tmpdir(), 'opensheet-xlsx-unicode-'));
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
try {
  writeFileSync(source, generated.bytes, { flag: 'wx' });
  const version = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const log = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${pathToFileURL(join(temp, 'profile')).href}`,
      '--headless',
      '--convert-to',
      'xlsx',
      '--outdir',
      temp,
      source,
    ],
    { encoding: 'utf8', timeout: 60000 },
  ).trim();
  const savedBytes = readFileSync(join(temp, basename(source)));
  writeFileSync(saved, savedBytes, { flag: 'wx' });
  const python = `import json,openpyxl,sys
book=openpyxl.load_workbook(sys.argv[1]); sheet=book['Unicode']
print(json.dumps({'version':openpyxl.__version__,'cells':[sheet[f'A{i}'].value for i in range(1,9)],'note':sheet['A5'].comment.text if sheet['A5'].comment else None,'link':sheet['A6'].hyperlink.target if sheet['A6'].hyperlink else None,'tooltips':[sheet[f'A{i}'].hyperlink.tooltip if sheet[f'A{i}'].hyperlink else None for i in (6,7,8)]},ensure_ascii=True))`;
  const inspect = (path) =>
    JSON.parse(execFileSync('python3', ['-c', python, path], { encoding: 'utf8', timeout: 30000 }));
  const before = inspect(source);
  const after = inspect(saved);
  if (before.cells.slice(0, 4).join('|') !== ['�', '�', '🙂', '_xfffe_'].join('|'))
    throw new Error('Unexpected independent source values');
  if (JSON.stringify(before.tooltips) !== JSON.stringify(['tip�', 'ordinary tip', 'tip🙂']))
    throw new Error('The input workbook did not contain the expected tooltip controls');
  const report = {
    reader: version,
    scope:
      'Synthetic OpenSheet adapter XLSX with unpaired surrogates, one paired emoji, U+FFFE, a note and a link tooltip. openpyxl inspects source and LibreOfficeDev-saved output; no Excel or stable LibreOffice guarantee.',
    inputs: cells,
    adapterReport: generated.report,
    artifacts: {
      source: { file: basename(source), sha256: sha256(generated.bytes) },
      saved: { file: basename(saved), sha256: sha256(savedBytes) },
    },
    versions: { node: process.version, openpyxl: before.version },
    source: {
      commit: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
      coreBundleSha256: sha256(readFileSync('packages/core/dist/index.js')),
      adapterBundleSha256: sha256(readFileSync('packages/adapter-sheetjs/dist/index.js')),
      adapterCellsSourceSha256: sha256(readFileSync('packages/adapter-sheetjs/src/cells.ts')),
      scriptSha256: sha256(readFileSync('scripts/xlsx-unicode-reader.mjs')),
      lockfileSha256: sha256(readFileSync('pnpm-lock.yaml')),
    },
    sourceInspection: before,
    savedInspection: after,
    conversionLog: log,
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(
    JSON.stringify({ reader: version, source: before, saved: after, report: reportPath }),
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
