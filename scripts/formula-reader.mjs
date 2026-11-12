import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  mkdtempSync,
  copyFileSync,
  existsSync,
  rmSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve, basename } from 'node:path';
import { transform } from 'esbuild';
import * as XLSX from 'xlsx';
import { toXlsxFormula } from '../packages/adapter-sheetjs/dist/index.js';

const outputArg = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!outputArg)
  throw new Error(
    'Pass --output=path/to/new-result.json; historical evidence is never overwritten',
  );
const output = resolve(outputArg);
if (existsSync(output)) throw new Error(`Evidence already exists: ${output}`);
const fixtureSource = readFileSync('tests/fixtures/formula-corpus.ts', 'utf8');
const { code } = await transform(fixtureSource, { loader: 'ts', format: 'esm' });
const corpusModule = await import(
  `data:text/javascript;base64,${Buffer.from(code).toString('base64')}`
);
const dateSystem = process.argv.includes('--date-system=1904') ? '1904' : '1900';
const formulaCorpus =
  dateSystem === '1904' ? corpusModule.date1904Corpus : corpusModule.formulaCorpus;
const probes = [
  'SUM("7",TRUE,2)',
  'AVERAGE("7",TRUE,2)',
  'MIN("7",TRUE,2)',
  'MAX("7",TRUE,2)',
  'AND("TRUE")',
  'OR("FALSE")',
];
const work = mkdtempSync(join(tmpdir(), 'opensheet-reader-'));
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
try {
  const version = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const source = join(work, 'semantics.xlsx');
  const converted = join(work, 'recalculated');
  mkdirSync(converted);
  const sheet = XLSX.utils.aoa_to_sheet([[true], ['7'], [2], [null], [''], [null], [false]]);
  XLSX.utils.sheet_add_aoa(sheet, corpusModule.formulaLookupInputs, { origin: 'D1' });
  XLSX.utils.sheet_add_aoa(sheet, corpusModule.orderedLookupInputs, { origin: 'G1' });
  sheet.A6 = { t: 'n', f: '1/0' };
  const expressions = [...formulaCorpus.map(([expression]) => expression), ...probes];
  expressions.forEach((expression, i) => {
    sheet[`B${i + 1}`] = { t: 'n', f: toXlsxFormula(expression) };
    sheet[`C${i + 1}`] = { t: 's', v: expression };
  });
  sheet['!ref'] = `A1:I${expressions.length}`;
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, 'Semantics');
  workbook.Workbook = {
    WBProps: { date1904: dateSystem === '1904' },
    CalcPr: { fullCalcOnLoad: true },
  };
  writeFileSync(source, XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }));
  const log = execFileSync(
    'soffice',
    ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', converted, source],
    { encoding: 'utf8', timeout: 120000 },
  );
  const recalculated = join(converted, basename(source));
  if (!existsSync(recalculated)) throw new Error(`Reader did not produce a workbook: ${log}`);
  const result = XLSX.read(readFileSync(recalculated), { type: 'buffer' }).Sheets.Semantics;
  const observations = expressions.map((expression, i) => {
    const cell = result[`B${i + 1}`];
    if (!cell || cell.v === undefined)
      throw new Error(`Missing independently calculated result at B${i + 1}`);
    const actual = cell.t === 'e' ? { error: cell.w } : cell.v;
    const fixture = formulaCorpus[i];
    return {
      expression,
      tag: fixture?.[2] ?? 'coercion-probe',
      expected: fixture?.[1],
      actual,
      agrees: fixture ? JSON.stringify(actual) === JSON.stringify(fixture[1]) : null,
      readerFormula: cell.f,
    };
  });
  mkdirSync(dirname(output), { recursive: true });
  const sourceEvidence = output.replace(/\.json$/, '') + '.source.xlsx';
  const recalculatedEvidence = output.replace(/\.json$/, '') + '.recalculated.xlsx';
  if (existsSync(sourceEvidence) || existsSync(recalculatedEvidence))
    throw new Error('Workbook evidence already exists');
  copyFileSync(source, sourceEvidence);
  copyFileSync(recalculated, recalculatedEvidence);
  const report = {
    schemaVersion: 1,
    dateSystem,
    date: new Date().toISOString(),
    reader: version,
    sourceCommit: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
    sourceDirty: Boolean(
      execFileSync('git', ['status', '--porcelain'], { encoding: 'utf8' }).trim(),
    ),
    provenance:
      'Synthetic public formula corpus; expected results hand-authored, cached actual results produced by LibreOffice conversion, decoded by SheetJS. No OpenSheet-generated expected values.',
    fixtureSha256: digest(fixtureSource),
    sourceWorkbook: basename(sourceEvidence),
    sourceSha256: digest(readFileSync(sourceEvidence)),
    recalculatedWorkbook: basename(recalculatedEvidence),
    recalculatedSha256: digest(readFileSync(recalculatedEvidence)),
    conversionLog: log.trim(),
    observations,
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
  process.stdout.write(
    JSON.stringify(
      {
        output,
        reader: version,
        fixtures: formulaCorpus.length,
        agreements: observations.filter((entry) => entry.agrees === true).length,
        differences: observations.filter((entry) => entry.agrees === false),
        probes: observations.filter((entry) => entry.agrees === null),
      },
      null,
      2,
    ) + '\n',
  );
} finally {
  rmSync(work, { recursive: true, force: true });
}
