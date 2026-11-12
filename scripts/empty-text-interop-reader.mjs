import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!output) throw new Error('Supply a new --output path');
const reportPath = resolve(output),
  prefix = reportPath.replace(/\.json$/, '');
const paths = { exported: prefix + '.xlsx', saved: prefix + '.saved.xlsx' };
if ([reportPath, ...Object.values(paths)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-empty-text-'));
try {
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 4, columns: 2 }] });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: {
        strings: { positions: [0, 32], values: ['', 'x'] },
        formulas: { positions: [1], values: ['COUNTA(A1:A2)'] },
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
  if (
    strictCode !== 'UNSUPPORTED_FEATURE' ||
    !exported.report.issues.some((issue) => issue.code === 'EMPTY_TEXT_INTEROP')
  )
    throw new Error('Missing empty-text compatibility warning');
  const source = join(work, 'source.xlsx');
  writeFileSync(source, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
  const profile = pathToFileURL(join(work, 'profile')).href;
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const savedDir = join(work, 'saved');
  execFileSync('mkdir', [savedDir]);
  const log = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${profile}`,
      '--headless',
      '--convert-to',
      'xlsx:Calc MS Excel 2007 XML',
      '--outdir',
      savedDir,
      source,
    ],
    { encoding: 'utf8', timeout: 120000 },
  ).trim();
  const saved = join(savedDir, 'source.xlsx');
  const read = (path) =>
    XLSX.read(readFileSync(path), { type: 'buffer', sheetStubs: true, xlfn: true });
  const before = read(source),
    after = read(saved);
  const describe = (book) => {
    const sheet = book.Sheets.Data;
    return {
      emptyCell: sheet.A1 ?? null,
      nonemptyCell: sheet.A2 ?? null,
      countFormula: sheet.B1 ?? null,
    };
  };
  const reimport = fromSheetJS(after);
  copyFileSync(source, paths.exported);
  copyFileSync(saved, paths.saved);
  const artifacts = Object.fromEntries(
    Object.entries(paths).map(([name, path]) => [
      name,
      {
        file: path.split('/').at(-1),
        sha256: createHash('sha256').update(readFileSync(path)).digest('hex'),
      },
    ]),
  );
  const result = {
    startedAt: new Date().toISOString(),
    reader,
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
    strictCode,
    exportReport: exported.report,
    before: describe(before),
    after: describe(after),
    reimportReport: reimport.report,
    artifacts,
    log,
  };
  writeFileSync(reportPath, JSON.stringify(result, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify({ reader, strictCode, before: result.before, after: result.after }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
