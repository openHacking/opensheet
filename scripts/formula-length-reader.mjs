import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync,
  constants,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { createWorkbookFile, encodeFileCells } from '../packages/core/dist/index.js';
import { writeXlsx } from '../packages/adapter-sheetjs/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!output?.endsWith('.json')) throw new Error('Supply a new --output=NAME.json');
const reportPath = resolve(output),
  prefix = reportPath.slice(0, -5);
const artifacts = { exported: `${prefix}.xlsx`, saved: `${prefix}.saved.xlsx` };
if ([reportPath, ...Object.values(artifacts)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-formula-length-'));
try {
  const sizes = [8189, 8190, 8191, 8192, 8193, 16383];
  const formulas = sizes.map(
    (size) => (size % 2 ? '1' : '+1') + '+0'.repeat(Math.floor((size - 1) / 2)),
  );
  const file = createWorkbookFile({ sheets: [{ name: 'Length', rows: sizes.length, columns: 1 }] });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        formulas.map((expression, rowId) => ({
          rowId,
          columnId: 0,
          input: { type: 'formula', expression },
        })),
        new Map(),
      ),
    },
  ];
  const exported = writeXlsx(file, XLSX);
  const source = join(work, 'source.xlsx'),
    savedDir = join(work, 'saved');
  writeFileSync(source, exported.bytes);
  mkdirSync(savedDir);
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
  const python = `import json,sys,openpyxl\nr={'version':openpyxl.__version__,'files':[]}\nfor p in sys.argv[1:]:\n f=openpyxl.load_workbook(p,data_only=False)['Length']; v=openpyxl.load_workbook(p,data_only=True)['Length']\n r['files'].append([{'address':f.cell(i,1).coordinate,'formulaLength':len(f.cell(i,1).value) if isinstance(f.cell(i,1).value,str) else None,'cached':v.cell(i,1).value,'cacheType':v.cell(i,1).data_type} for i in range(1,7)])\nprint(json.dumps(r))`;
  const independent = JSON.parse(
    execFileSync('python3', ['-c', python, source, saved], { encoding: 'utf8', timeout: 30000 }),
  );
  for (const [name, path] of Object.entries({ exported: source, saved }))
    copyFileSync(path, artifacts[name], constants.COPYFILE_EXCL);
  const report = {
    date: new Date().toISOString(),
    provenance:
      'Synthetic addition formulas exported by OpenSheet, recalculated and saved by LibreOfficeDev, then inspected with openpyxl. Other formula shapes, Excel and visual behavior are untested.',
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
    sizesExcludingEquals: sizes,
    reader,
    independent,
    exportReport: exported.report,
    log,
    artifacts: Object.fromEntries(
      Object.entries(artifacts).map(([name, path]) => [
        name,
        {
          file: basename(path),
          sha256: createHash('sha256').update(readFileSync(path)).digest('hex'),
        },
      ]),
    ),
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify({ reader, independent }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
