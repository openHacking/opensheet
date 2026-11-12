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
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const arg = process.argv.find((value) => value.startsWith('--output='))?.slice(9);
if (!arg?.endsWith('.json')) throw new Error('Pass --output=NEW_REPORT.json');
const output = resolve(arg),
  prefix = output.slice(0, -5);
const sourceTarget = `${prefix}.source.xlsx`,
  savedTarget = `${prefix}.recalculated.xlsx`;
if ([output, sourceTarget, savedTarget].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-xlsx-limit-'));
const hash = (data) => createHash('sha256').update(data).digest('hex');
try {
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const sheet = XLSX.utils.aoa_to_sheet([[1]]);
  for (let row = 2; row <= 130; row++) sheet[`A${row}`] = { t: 'n', f: `A${row - 1}+1` };
  sheet['!ref'] = 'A1:A130';
  const input = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(input, sheet, 'Data');
  const file = fromSheetJS(input).file;
  const exported = toSheetJS(file);
  const uncalculated = exported.workbook.Sheets.Data;
  for (let row = 2; row <= 130; row++)
    if (uncalculated[`A${row}`]?.f !== `A${row - 1}+1` || uncalculated[`A${row}`]?.v !== undefined)
      throw new Error(`Incorrect formula or cache at ${row}`);
  const source = join(work, 'chain.xlsx'),
    target = join(work, 'out');
  mkdirSync(target);
  writeFileSync(source, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
  const written = XLSX.read(readFileSync(source), {
    type: 'buffer',
    cellFormula: true,
    sheetStubs: true,
  }).Sheets.Data;
  const writtenFormulas = Array.from(
    { length: 129 },
    (_, index) => written[`A${index + 2}`]?.f ?? null,
  );
  if (writtenFormulas.some((formula, index) => formula !== `A${index + 1}+1`))
    throw new Error('Serialized formula missing');
  const log = execFileSync(
    'soffice',
    ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', target, source],
    { encoding: 'utf8', timeout: 120000 },
  );
  const saved = join(target, 'chain.xlsx');
  if (!existsSync(saved)) throw new Error(`Missing recalculated file: ${log}`);
  const result = XLSX.read(readFileSync(saved), { type: 'buffer', cellFormula: true }).Sheets.Data;
  const observations = Array.from({ length: 130 }, (_, index) => ({
    address: `A${index + 1}`,
    value: result[`A${index + 1}`]?.v ?? null,
    formula: result[`A${index + 1}`]?.f ?? null,
  }));
  const agrees = observations.every((cell, index) => cell.value === index + 1);
  mkdirSync(dirname(output), { recursive: true });
  copyFileSync(source, sourceTarget);
  copyFileSync(saved, savedTarget);
  writeFileSync(
    output,
    JSON.stringify(
      {
        reader,
        node: process.version,
        source: sourceFingerprint(),
        adapterBuild: localBuildFingerprint('packages/adapter-sheetjs/dist'),
        report: exported.report,
        observations,
        agrees,
        log,
        artifacts: [sourceTarget, savedTarget].map((path) => ({
          name: basename(path),
          sha256: hash(readFileSync(path)),
        })),
        limitations:
          'One synthetic 130-cell addition chain; XLSX emitted by SheetJS, recalculated by LibreOfficeDev headless and decoded by SheetJS. Does not establish Excel, native clipboard, other formulas or general cache/recalculation behavior.',
      },
      null,
      2,
    ) + '\n',
    { flag: 'wx' },
  );
  console.log(
    JSON.stringify({ reader, agrees, last: observations.at(-1), issues: exported.report.issues }),
  );
} finally {
  rmSync(work, { recursive: true, force: true });
}
