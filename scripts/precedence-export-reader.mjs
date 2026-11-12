import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  constants,
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import * as XLSX from 'xlsx';
import { createWorkbookFile, encodeFileCells } from '../packages/core/dist/index.js';
import { toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';

const output = resolve(process.argv[2] ?? '');
if (!process.argv[2])
  throw new Error('Usage: node scripts/precedence-export-reader.mjs NEW_REPORT');
const sourcePath = output.replace(/\.json$/, '.source.xlsx');
const readerPath = output.replace(/\.json$/, '.recalculated.xlsx');
if (
  new Set([output, sourcePath, readerPath]).size !== 3 ||
  [output, sourcePath, readerPath].some(existsSync)
)
  throw new Error('Use a new .json report path');
const expressions = ['2^3^2', '-2^2', '2^(-3^2)', '(-2)^2'];
const expected = [512, -4, 1 / 512, 4];
const file = createWorkbookFile({ sheets: [{ name: 'Power' }] });
file.sheets[0].blocks = [
  {
    row: 0,
    column: 0,
    data: encodeFileCells(
      expressions.map((expression, rowId) => ({
        rowId,
        columnId: 0,
        input: { type: 'formula', expression },
      })),
      new Map(),
    ),
  },
];
const converted = toSheetJS(file);
const source = XLSX.write(converted.workbook, { type: 'buffer', bookType: 'xlsx' });
const temporary = mkdtempSync(join(tmpdir(), 'opensheet-precedence-export-'));
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
try {
  const local = join(temporary, 'input.xlsx');
  const resultDirectory = join(temporary, 'result');
  mkdirSync(resultDirectory);
  writeFileSync(local, source);
  const version = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const log = execFileSync(
    'soffice',
    [
      '--headless',
      '--convert-to',
      'xlsx:Calc MS Excel 2007 XML',
      '--outdir',
      resultDirectory,
      local,
    ],
    { encoding: 'utf8', timeout: 120000 },
  );
  const resultPath = join(resultDirectory, basename(local));
  const reader = readFileSync(resultPath);
  const result = XLSX.read(reader, { type: 'buffer' }).Sheets.Power;
  const actual = expressions.map((_, index) => result[`A${index + 1}`]?.v);
  if (actual.some((value, index) => value !== expected[index]))
    throw new Error(`Independent result mismatch: ${JSON.stringify(actual)}`);
  writeFileSync(sourcePath, source, { flag: 'wx' });
  copyFileSync(resultPath, readerPath, constants.COPYFILE_EXCL);
  writeFileSync(
    output,
    JSON.stringify(
      {
        reader: version,
        expressions,
        expected,
        actual,
        exported: expressions.map((_, index) => converted.workbook.Sheets.Power[`A${index + 1}`].f),
        issues: converted.report.issues,
        sourceSha256: digest(source),
        resultSha256: digest(reader),
        log,
        scope:
          'Four synthetic native formulas exported through the adapter and independently recalculated by a LibreOffice development build; not Excel or general formula compatibility.',
      },
      null,
      2,
    ) + '\n',
    { flag: 'wx' },
  );
} finally {
  rmSync(temporary, { recursive: true, force: true });
}
