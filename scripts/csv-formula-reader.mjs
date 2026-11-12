import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import * as XLSX from 'xlsx';
import { createWorkbookFile } from '../packages/core/dist/index.js';
import { exportRange } from '../packages/formats/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const arg = process.argv.find((value) => value.startsWith('--output='))?.slice(9);
if (!arg || !arg.endsWith('.json')) throw new Error('Pass --output=NEW_REPORT.json');
const output = resolve(arg),
  prefix = output.slice(0, -5);
const targets = [output, `${prefix}.csv`, `${prefix}.xlsx`];
if (targets.some(existsSync)) throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-csv-formula-'));
const hash = (data) => createHash('sha256').update(data).digest('hex');
try {
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 1, columns: 5 }] });
  file.sheets[0].blocks = [
    {
      row: 0,
      column: 0,
      data: {
        numbers: { positions: [0], values: [5] },
        formulas: { positions: [1, 2, 4], values: ['A1+3', 'SUM(A1,2)', 'IF(A1>0,"yes","no")'] },
        strings: { positions: [3], values: ['=2+3'] },
      },
    },
  ];
  const exported = exportRange(file, {
    sheetId: file.sheets[0].id,
    range: 'A1:E1',
    format: 'csv',
    options: { safe: false, formulas: 'preserve' },
  });
  const input = join(work, 'formula.csv'),
    target = join(work, 'out');
  mkdirSync(target);
  writeFileSync(input, exported.text);
  const log = execFileSync(
    'soffice',
    ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', target, input],
    { encoding: 'utf8', timeout: 120000 },
  );
  const converted = join(target, 'formula.xlsx');
  if (!existsSync(converted)) throw new Error(`Missing reader output: ${log}`);
  const workbook = XLSX.read(readFileSync(converted), { type: 'buffer', cellFormula: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const cells = ['A1', 'B1', 'C1', 'D1', 'E1'].map((address) => ({
    address,
    type: sheet[address]?.t ?? null,
    value: sheet[address]?.v ?? null,
    formula: sheet[address]?.f ?? null,
  }));
  mkdirSync(dirname(output), { recursive: true });
  copyFileSync(input, targets[1]);
  copyFileSync(converted, targets[2]);
  writeFileSync(
    output,
    JSON.stringify(
      {
        reader,
        node: process.version,
        sourceFingerprint: sourceFingerprint(),
        formatsBuild: localBuildFingerprint('packages/formats/dist'),
        report: exported.report,
        cells,
        log,
        artifacts: targets
          .slice(1)
          .map((path) => ({ name: basename(path), sha256: hash(readFileSync(path)) })),
        limitations:
          'Synthetic benign formulas and one formula-like text cell; LibreOfficeDev headless defaults, XLSX decoded with SheetJS. No Excel or locale compatibility claim.',
      },
      null,
      2,
    ) + '\n',
    { flag: 'wx' },
  );
  console.log(JSON.stringify({ reader, report: exported.report, cells }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
