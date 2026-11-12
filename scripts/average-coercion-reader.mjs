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
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { evaluateFormula, parseFormula } from '../packages/formula/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const arg = process.argv.find((value) => value.startsWith('--output='))?.slice(9);
if (!arg?.endsWith('.json')) throw new Error('Pass --output=NEW_REPORT.json');
const output = resolve(arg),
  prefix = output.slice(0, -5),
  targets = [`${prefix}.source.xlsx`, `${prefix}.recalculated.xlsx`];
if ([output, ...targets].some(existsSync)) throw new Error('Refusing to overwrite evidence');
const formulas = [
  'AVERAGE(TRUE,2)',
  'AVERAGE(FALSE,2)',
  'AVERAGE("7",2)',
  'AVERAGE("bad",2)',
  'AVERAGE("",2)',
  'AVERAGE(A1,2)',
  'AVERAGE(A2,2)',
  'AVERAGE(A3,2)',
];
const inputs = [true, '7', 0];
const work = mkdtempSync(join(tmpdir(), 'opensheet-average-reader-'));
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
try {
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const sheet = XLSX.utils.aoa_to_sheet(inputs.map((value) => [value]));
  formulas.forEach((formula, index) => {
    sheet[`B${index + 1}`] = { t: 'n', f: formula };
  });
  sheet['!ref'] = `A1:B${formulas.length}`;
  const input = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(input, sheet, 'Data');
  const exported = toSheetJS(fromSheetJS(input).file);
  const source = join(work, 'average.xlsx'),
    out = join(work, 'out');
  mkdirSync(out);
  writeFileSync(source, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
  const log = execFileSync(
    'soffice',
    ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', out, source],
    { encoding: 'utf8', timeout: 120000 },
  );
  const saved = join(out, 'average.xlsx');
  if (!existsSync(saved)) throw new Error(`Missing output: ${log}`);
  const result = XLSX.read(readFileSync(saved), { type: 'buffer', cellFormula: true }).Sheets.Data;
  const observations = formulas.map((formula, index) => {
    const cell = result[`B${index + 1}`];
    const actual =
      !cell || cell.v === undefined ? null : cell.t === 'e' ? { error: cell.w } : cell.v;
    const openSheet = evaluateFormula(parseFormula(formula), (ref) =>
      ref.column === 0 ? (inputs[ref.row] ?? null) : null,
    );
    return {
      formula,
      openSheet,
      actual,
      readerFormula: cell?.f ?? null,
      agrees: JSON.stringify(openSheet) === JSON.stringify(actual),
    };
  });
  mkdirSync(dirname(output), { recursive: true });
  copyFileSync(source, targets[0]);
  copyFileSync(saved, targets[1]);
  writeFileSync(
    output,
    JSON.stringify(
      {
        reader,
        node: process.version,
        source: sourceFingerprint(),
        adapterBuild: localBuildFingerprint('packages/adapter-sheetjs/dist'),
        formulaBuild: localBuildFingerprint('packages/formula/dist'),
        observations,
        log,
        artifacts: targets.map((path) => ({
          name: basename(path),
          sha256: hash(readFileSync(path)),
        })),
        limitations:
          'Synthetic single-cell scalar cases; LibreOfficeDev headless, XLSX generated and decoded with SheetJS. No actual Excel, locale or general formula compatibility claim.',
      },
      null,
      2,
    ) + '\n',
    { flag: 'wx' },
  );
  console.log(JSON.stringify({ reader, observations }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
