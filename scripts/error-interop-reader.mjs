import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  mkdtempSync,
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
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { toSheetJS, fromSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { parseDelimited } from '../packages/formats/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!output) throw new Error('Supply a new --output path');
const reportPath = resolve(output),
  prefix = reportPath.replace(/\.json$/, '');
const artifacts = {
  exported: prefix + '.xlsx',
  csv: prefix + '.csv',
  saved: prefix + '.recalculated.xlsx',
};
if ([reportPath, ...Object.values(artifacts)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-errors-'));
const file = createWorkbookFile({ sheets: [{ name: 'Errors', rows: 4, columns: 1 }] });
file.sheets[0].blocks = [
  {
    row: 0,
    column: 0,
    data: {
      errors: { positions: [0, 32, 64, 96], values: ['#DIV/0!', '#N/A', '#CYCLE!', '#SPILL!'] },
    },
  },
];
const exported = toSheetJS(file);
let strictCode = '';
try {
  toSheetJS(file, { unsupported: 'strict' });
} catch (error) {
  strictCode = error.code;
}
if (strictCode !== 'UNSUPPORTED_FEATURE')
  throw new Error('Strict export failed to reject unsupported errors');
const source = join(work, 'source.xlsx');
writeFileSync(source, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
const profile = pathToFileURL(join(work, 'profile')).href;
const reader = execFileSync('soffice', ['--version'], { encoding: 'utf8', timeout: 30000 }).trim();
try {
  const log = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${profile}`,
      '--headless',
      '--convert-to',
      'csv:Text - txt - csv (StarCalc):44,34,76,1',
      '--outdir',
      work,
      source,
    ],
    { encoding: 'utf8', timeout: 120000 },
  ).trim();
  const csv = readFileSync(join(work, 'source.csv'), 'utf8');
  const saveDir = join(work, 'save');
  execFileSync('mkdir', [saveDir]);
  const saveLog = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${profile}`,
      '--headless',
      '--convert-to',
      'xlsx:Calc MS Excel 2007 XML',
      '--outdir',
      saveDir,
      source,
    ],
    { encoding: 'utf8', timeout: 120000 },
  ).trim();
  const savedPath = join(saveDir, 'source.xlsx');
  const bytes = readFileSync(savedPath);
  const decoded = XLSX.read(bytes, { type: 'buffer' });
  const imported = fromSheetJS(decoded);
  const observations = ['A1', 'A2', 'A3', 'A4'].map((cell) => ({
    cell,
    type: decoded.Sheets.Errors[cell]?.t,
    value: decoded.Sheets.Errors[cell]?.v,
    text: decoded.Sheets.Errors[cell]?.w,
  }));
  const hash = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');
  const expectedDisplay = ['#DIV/0!', '#N/A', '#VALUE!', '#VALUE!'];
  const csvRows = parseDelimited(csv);
  const assertions = expectedDisplay.map((expected, i) => ({
    cell: `A${i + 1}`,
    expected,
    csv: csvRows[i]?.[0] ?? null,
    saved: observations[i].text ?? null,
    agrees: csvRows[i]?.[0] === expected && observations[i].text === expected,
  }));
  copyFileSync(source, artifacts.exported, constants.COPYFILE_EXCL);
  copyFileSync(join(work, 'source.csv'), artifacts.csv, constants.COPYFILE_EXCL);
  copyFileSync(savedPath, artifacts.saved, constants.COPYFILE_EXCL);
  const report = {
    date: new Date().toISOString(),
    reader,
    provenance:
      'CC0 synthetic OpenSheet error cells, exported with SheetJS adapter; LibreOffice converted the XLSX to CSV and independently saved to XLSX. CSV is the reader display observation; the saved XLSX is decoded with SheetJS.',
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
    inputErrors: ['#DIV/0!', '#N/A', '#CYCLE!', '#SPILL!'],
    expectedDisplay,
    assertions,
    strictCode,
    exportReport: exported.report,
    csv,
    observations,
    importReport: imported.report,
    logs: [log, saveLog],
    artifacts: Object.fromEntries(
      Object.entries(artifacts).map(([key, value]) => [
        key,
        { file: basename(value), sha256: hash(value) },
      ]),
    ),
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  if (assertions.some((assertion) => !assertion.agrees))
    throw new Error('Reader output differs from expected errors');
  process.stdout.write(
    JSON.stringify(
      { reader, csv, observations, strictCode, issues: exported.report.issues },
      null,
      2,
    ) + '\n',
  );
} finally {
  rmSync(work, { recursive: true, force: true });
}
