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
if (!output || !output.endsWith('.json')) throw new Error('Supply a new --output=NAME.json');
const reportPath = resolve(output),
  prefix = reportPath.slice(0, -5);
const artifacts = { exported: `${prefix}.xlsx`, saved: `${prefix}.saved.xlsx` };
if ([reportPath, ...Object.values(artifacts)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-number-precedence-'));
try {
  const file = createWorkbookFile({ sheets: [{ name: 'Formats', rows: 3, columns: 1 }] });
  file.styles = [{ numberFormat: '0.00' }];
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: encodeFileCells(
        [
          {
            rowId: 0,
            columnId: 0,
            input: { type: 'number', value: 0.25 },
            styleId: 's0',
            numberFormat: '0%',
          },
          { rowId: 1, columnId: 0, input: { type: 'number', value: 0.25 }, styleId: 's0' },
          {
            rowId: 2,
            columnId: 0,
            input: { type: 'blank' },
            styleId: 's0',
            link: { target: 'https://example.com/' },
          },
        ],
        new Map([['s0', 0]]),
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
  const python = `import json,sys,openpyxl\nr={'version':openpyxl.__version__,'files':[]}\nfor path in sys.argv[1:]:\n s=openpyxl.load_workbook(path)['Formats']\n r['files'].append([{'address':c.coordinate,'value':c.value,'format':c.number_format,'link':c.hyperlink.target if c.hyperlink else None} for c in (s['A1'],s['A2'],s['A3'])])\nprint(json.dumps(r))`;
  const independent = JSON.parse(
    execFileSync('python3', ['-c', python, source, saved], { encoding: 'utf8', timeout: 30000 }),
  );
  for (const [name, path] of Object.entries({ exported: source, saved }))
    copyFileSync(path, artifacts[name], constants.COPYFILE_EXCL);
  const result = {
    date: new Date().toISOString(),
    provenance:
      'Synthetic cells exported by OpenSheet; LibreOffice saved the XLSX and openpyxl inspected both files. Excel and visual rendering are not tested.',
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
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
  writeFileSync(reportPath, JSON.stringify(result, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify({ reader, independent }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
