import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { FileReader } from '../packages/core/dist/index.js';
import { fromSheetJS } from '../packages/adapter-sheetjs/dist/index.js';

const prefix = process.argv[2];
if (!prefix) throw new Error('Usage: node scripts/xlsx-range-link-reader.mjs NEW_PREFIX');
const source = resolve(`${prefix}.xlsx`),
  saved = resolve(`${prefix}.saved.xlsx`),
  reportPath = resolve(`${prefix}.json`);
const temp = mkdtempSync(join(tmpdir(), 'opensheet-range-links-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const python = (action, path) =>
  execFileSync('python3', ['scripts/xlsx-range-link-fixture.py', action, path], {
    encoding: 'utf8',
    timeout: 30000,
  });
const inspect = (path) => JSON.parse(python('inspect', path));
const imported = (path) => {
  const book = XLSX.read(readFileSync(path), { type: 'buffer', bookFiles: true, cellHTML: false });
  const result = fromSheetJS(book);
  const reader = new FileReader(result.file);
  try {
    return {
      targets: Array.from(
        { length: 3 },
        (_, row) => reader.getCell(result.file.sheets[0].id, row, 0)?.link?.target,
      ),
      issues: result.report.issues,
    };
  } finally {
    reader.dispose();
  }
};
try {
  python('create', source);
  const before = inspect(source);
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
  writeFileSync(saved, readFileSync(join(temp, basename(source))), { flag: 'wx' });
  const after = inspect(saved);
  const sourceImport = imported(source),
    savedImport = imported(saved);
  const target = 'https://example.com/路径#片段';
  const normalized = (cell) => cell.target + (cell.location ? `#${cell.location}` : '');
  if (
    JSON.stringify(before.rawRefs) !== JSON.stringify(['A1:A3']) ||
    JSON.stringify(after.rawRefs) !== JSON.stringify(['A1', 'A2', 'A3'])
  )
    throw new Error('Unexpected hyperlink references');
  for (const inspection of [before, after]) {
    if (
      JSON.stringify(inspection.cells.map((cell) => cell.value)) !==
        JSON.stringify(['one', 'two', 'three']) ||
      inspection.cells.some((cell) => normalized(cell) !== target)
    )
      throw new Error('Independent readback differs');
  }
  for (const result of [sourceImport, savedImport]) {
    if (
      result.targets.some((value) => value !== target) ||
      !result.issues.some(
        (issue) => issue.code === 'XLSX_LINK_URL_UTF8_RESTORED' && issue.count === 3,
      )
    )
      throw new Error('Adapter failed to restore links');
  }
  const report = {
    reader: version,
    scope:
      'Synthetic openpyxl workbook with one manually edited OOXML range reference; openpyxl readback before and after one LibreOfficeDev headless save, followed by SheetJS/OpenSheet import. No Excel, stable LibreOffice, native navigation, or visual assessment.',
    before,
    after,
    sourceImport,
    savedImport,
    artifacts: Object.fromEntries(
      [
        ['source', source],
        ['saved', saved],
      ].map(([key, path]) => [key, { file: basename(path), sha256: hash(readFileSync(path)) }]),
    ),
    versions: { node: process.version, openpyxl: before.openpyxl },
    source: {
      commit: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
      adapterBundleSha256: hash(readFileSync('packages/adapter-sheetjs/dist/index.js')),
      fixtureScriptSha256: hash(readFileSync('scripts/xlsx-range-link-fixture.py')),
      readerScriptSha256: hash(readFileSync('scripts/xlsx-range-link-reader.mjs')),
    },
    conversionLog: log,
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(
    JSON.stringify({
      before,
      after,
      sourceTargets: sourceImport.targets,
      savedTargets: savedImport.targets,
    }),
  );
} finally {
  rmSync(temp, { recursive: true, force: true });
}
