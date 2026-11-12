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
import { basename, dirname, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import * as XLSX from 'xlsx';
import { stringifyDelimited } from '../packages/formats/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = resolve(process.argv.find((arg) => arg.startsWith('--output='))?.slice(9) ?? '');
if (!process.argv.some((arg) => arg.startsWith('--output=')) || !output.endsWith('.json'))
  throw new Error('Pass --output=NEW_REPORT.json');
if (existsSync(output)) throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-csv-reader-'));
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const samples = [
  ['equals', '=1+1'],
  ['plus', '+1'],
  ['minus', '-1'],
  ['at', '@SUM(1;1)'],
  ['leading-space', '  =1+1'],
  ['tab', '\ttext'],
  ['cr', '\rtext'],
  ['lf', '\ntext'],
  ['nul', '\0text'],
  ['fullwidth-equals', '＝1+1'],
  ['fullwidth-plus', '＋1'],
  ['fullwidth-minus', '－1'],
  ['fullwidth-at', '＠SUM(1;1)'],
  ['separator-and-quote', 'plain,"=1+1'],
  ['number', -5],
];
try {
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const observations = [];
  const outputs = [];
  mkdirSync(dirname(output), { recursive: true });
  for (const safe of [true, false]) {
    const label = safe ? 'safe' : 'raw';
    const source = join(work, `${label}.csv`);
    const firstDir = join(work, `${label}-first`),
      csvDir = join(work, `${label}-saved`),
      secondDir = join(work, `${label}-reopened`);
    for (const dir of [firstDir, csvDir, secondDir]) mkdirSync(dir);
    const encoded = stringifyDelimited(samples, ',', safe);
    writeFileSync(source, encoded.text);
    const convert = (input, dir, type) =>
      execFileSync('soffice', ['--headless', '--convert-to', type, '--outdir', dir, input], {
        encoding: 'utf8',
        timeout: 120000,
      });
    const logs = [convert(source, firstDir, 'xlsx:Calc MS Excel 2007 XML')];
    const first = join(firstDir, `${label}.xlsx`);
    if (!existsSync(first)) throw new Error(`Missing first workbook: ${logs[0]}`);
    logs.push(convert(first, csvDir, 'csv:Text - txt - csv (StarCalc):44,34,76,1'));
    const saved = join(csvDir, `${label}.csv`);
    if (!existsSync(saved)) throw new Error(`Missing saved CSV: ${logs[1]}`);
    logs.push(convert(saved, secondDir, 'xlsx:Calc MS Excel 2007 XML'));
    const second = join(secondDir, `${label}.xlsx`);
    if (!existsSync(second)) throw new Error(`Missing reopened workbook: ${logs[2]}`);
    const read = (path) => {
      const book = XLSX.read(readFileSync(path), { type: 'buffer', cellFormula: true });
      const sheet = book.Sheets[book.SheetNames[0]];
      return samples.map((sample, i) => {
        const labelCell = sheet[`A${i + 1}`],
          value = sheet[`B${i + 1}`];
        return {
          expectedLabel: sample[0],
          actualLabel: labelCell?.v ?? null,
          type: value?.t ?? null,
          value: value?.v ?? null,
          formula: value?.f ?? null,
        };
      });
    };
    observations.push({
      mode: label,
      sanitized: encoded.sanitized,
      first: read(first),
      reopened: read(second),
      logs,
    });
    for (const [stage, path] of [
      ['source.csv', source],
      ['first.xlsx', first],
      ['saved.csv', saved],
      ['reopened.xlsx', second],
    ]) {
      const destination = output.replace(/\.json$/, '') + `.${label}.${stage}`;
      if (existsSync(destination)) throw new Error(`Evidence exists: ${destination}`);
      copyFileSync(path, destination);
      outputs.push({ name: basename(destination), sha256: digest(readFileSync(destination)) });
    }
  }
  writeFileSync(
    output,
    JSON.stringify(
      {
        reader,
        node: process.version,
        sourceFingerprint: sourceFingerprint(),
        formatsBuild: localBuildFingerprint('packages/formats/dist'),
        source: 'Synthetic benign formula and text examples; no network formulas.',
        limitations:
          'One local LibreOffice build and its headless CSV defaults; SheetJS decodes saved XLSX. No Excel, other locales, UI import settings, or security guarantee.',
        samples,
        observations,
        outputs,
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
