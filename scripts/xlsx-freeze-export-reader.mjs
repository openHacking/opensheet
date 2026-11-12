import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { fromSheetJS, writeXlsx } from '../packages/adapter-sheetjs/dist/index.js';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
const startedAt = new Date().toISOString();
if (!output || !output.endsWith('.json')) throw new Error('Supply a new --output=NAME.json');
const reportPath = resolve(output),
  prefix = reportPath.slice(0, -5);
const paths = { exported: `${prefix}.xlsx`, saved: `${prefix}.saved.xlsx` };
if ([reportPath, ...Object.values(paths)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-freeze-export-'));
try {
  const definitions = [
    { name: 'Rows', rows: 4, columns: 0, cell: 'A5' },
    { name: 'Columns', rows: 0, columns: 3, cell: 'D1' },
    { name: "Both '中' &", rows: 3, columns: 2, cell: 'C4' },
    { name: 'Unfrozen', rows: 0, columns: 0, cell: null },
  ];
  const file = createWorkbookFile({
    sheets: definitions.map(({ name }) => ({ name, rows: 20, columns: 10 })),
  });
  for (let i = 0; i < definitions.length; i++) {
    file.sheets[i].freeze = { rows: definitions[i].rows, columns: definitions[i].columns };
    file.sheets[i].blocks = [
      { row: 0, column: 0, data: { numbers: { positions: [0], values: [100 + i] } } },
    ];
  }
  const exported = writeXlsx(file, XLSX, { unsupported: 'strict' });
  if (exported.report.issues.length) throw new Error('Unexpected strict export issues');
  const source = join(work, 'source.xlsx');
  writeFileSync(source, exported.bytes);
  const savedDir = join(work, 'saved');
  execFileSync('mkdir', [savedDir]);
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const log = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${pathToFileURL(join(work, 'profile')).href}`,
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
  const python = `import json,sys,openpyxl\nresult={'version':openpyxl.__version__,'files':[]}\nfor path in sys.argv[1:]:\n b=openpyxl.load_workbook(path, data_only=False)\n result['files'].append([{'name':s.title,'freeze':s.freeze_panes,'value':s['A1'].value} for s in b.worksheets])\nprint(json.dumps(result,ensure_ascii=False))`;
  const independent = JSON.parse(
    execFileSync('python3', ['-c', python, source, saved], { encoding: 'utf8', timeout: 30000 }),
  );
  const read = (path) =>
    fromSheetJS(
      XLSX.read(readFileSync(path), { type: 'buffer', bookFiles: true, sheetStubs: true }),
    );
  const imported = { exported: read(source), saved: read(saved) };
  for (const [name, path] of Object.entries({ exported: source, saved }))
    writeFileSync(paths[name], readFileSync(path), { flag: 'wx' });
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
    startedAt,
    source: sourceFingerprint(),
    reader,
    independent,
    expected: definitions.map(({ name, rows, columns, cell }, index) => ({
      name,
      freeze: cell,
      rows,
      columns,
      value: 100 + index,
    })),
    exportReport: exported.report,
    imported: Object.fromEntries(
      Object.entries(imported).map(([name, value]) => [
        name,
        {
          freezes: value.file.sheets.map((sheet) => ({ name: sheet.name, ...sheet.freeze })),
          report: value.report,
        },
      ]),
    ),
    artifacts,
    log,
  };
  writeFileSync(reportPath, JSON.stringify(result, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify({ reader, independent, imported: result.imported }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
