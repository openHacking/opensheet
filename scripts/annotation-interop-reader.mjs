import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  mkdtempSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
  existsSync,
  rmSync,
  copyFileSync,
  constants,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { createWorkbookFile, encodeFileCells, FileReader } from '../packages/core/dist/index.js';
import { toSheetJS, fromSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!output) throw new Error('Supply a new --output path');
const reportPath = resolve(output),
  prefix = reportPath.replace(/\.json$/, '');
const artifacts = { exported: prefix + '.xlsx', saved: prefix + '.reader.xlsx' };
if ([reportPath, ...Object.values(artifacts)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-annotations-'));
const file = createWorkbookFile({ sheets: [{ name: 'Annotations', rows: 7, columns: 1 }] });
const cells = [
  { rowId: 0, columnId: 0, input: { type: 'string', value: 'value' }, note: '' },
  { rowId: 1, columnId: 0, input: { type: 'blank' }, note: '' },
  { rowId: 2, columnId: 0, input: { type: 'string', value: 'value' }, note: 'hello' },
  { rowId: 3, columnId: 0, input: { type: 'blank' }, note: 'hello' },
  {
    rowId: 4,
    columnId: 0,
    input: { type: 'string', value: 'link' },
    link: { target: 'https://example.com/', tooltip: '' },
  },
  { rowId: 5, columnId: 0, input: { type: 'blank' }, link: { target: 'https://example.com/' } },
  {
    rowId: 6,
    columnId: 0,
    input: { type: 'blank' },
    note: 'blank with link',
    link: { target: 'https://example.com/' },
  },
];
file.sheets[0].blocks = [{ row: 0, column: 0, data: encodeFileCells(cells, new Map()) }];
const exported = toSheetJS(file);
let strictCode = '';
try {
  toSheetJS(file, { unsupported: 'strict' });
} catch (error) {
  strictCode = error.code;
}
if (strictCode !== 'UNSUPPORTED_FEATURE')
  throw new Error('Strict export did not report empty tooltip loss');
const source = join(work, 'source.xlsx');
writeFileSync(source, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
const reader = execFileSync('soffice', ['--version'], { encoding: 'utf8', timeout: 30000 }).trim();
try {
  const saveDir = join(work, 'save');
  mkdirSync(saveDir);
  const log = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${pathToFileURL(join(work, 'profile')).href}`,
      '--headless',
      '--convert-to',
      'xlsx:Calc MS Excel 2007 XML',
      '--outdir',
      saveDir,
      source,
    ],
    { encoding: 'utf8', timeout: 120000 },
  ).trim();
  const saved = join(saveDir, 'source.xlsx');
  function inspect(path) {
    const raw = XLSX.read(readFileSync(path), { type: 'buffer', sheetStubs: true });
    const imported = fromSheetJS(raw);
    const book = new FileReader(imported.file),
      id = imported.file.sheets[0].id;
    const result = cells.map((_, row) => {
      const c = book.getCell(id, row, 0);
      return {
        cell: `A${row + 1}`,
        input: c?.input ?? { type: 'blank' },
        note: c?.note ?? null,
        link: c?.link ?? null,
      };
    });
    book.dispose();
    return result;
  }
  const emitted = inspect(source),
    reopened = inspect(saved);
  const comparisons = emitted.map((entry, i) => ({
    cell: entry.cell,
    agrees: JSON.stringify(entry) === JSON.stringify(reopened[i]),
    emitted: entry,
    reopened: reopened[i],
  }));
  copyFileSync(source, artifacts.exported, constants.COPYFILE_EXCL);
  copyFileSync(saved, artifacts.saved, constants.COPYFILE_EXCL);
  const report = {
    date: new Date().toISOString(),
    reader,
    provenance:
      'CC0 synthetic OpenSheet annotation cells exported via SheetJS; LibreOffice opened and saved XLSX. Observations reimported through SheetJS; no visual inspection.',
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
    strictCode,
    exportReport: exported.report,
    comparisons,
    log,
    artifacts: Object.fromEntries(
      Object.entries(artifacts).map(([key, path]) => [
        key,
        {
          file: basename(path),
          sha256: createHash('sha256').update(readFileSync(path)).digest('hex'),
        },
      ]),
    ),
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  process.stdout.write(
    JSON.stringify({ reader, comparisons, issues: exported.report.issues }, null, 2) + '\n',
  );
} finally {
  rmSync(work, { recursive: true, force: true });
}
