import { build } from 'esbuild';
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
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const outputArg = process.argv.find((argument) => argument.startsWith('--output='))?.slice(9);
if (!outputArg?.endsWith('.json')) throw new Error('Pass --output=NEW_REPORT.json');
const output = resolve(outputArg),
  prefix = output.slice(0, -5);
const paths = ['blank', 'control'].flatMap((name) =>
  ['tsv', 'xlsx', 'csv'].map((ext) => `${prefix}-${name}.${ext}`),
);
if ([output, ...paths].some(existsSync)) throw new Error('Refusing to overwrite evidence');
const bundled = await build({
  entryPoints: ['packages/renderer/src/clipboard-text.ts'],
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
});
const { encodeClipboardText } = await import(
  `data:text/javascript;base64,${Buffer.from(bundled.outputFiles[0].text).toString('base64')}`
);
const work = mkdtempSync(join(tmpdir(), 'opensheet-clipboard-reader-'));
const hash = (buffer) => createHash('sha256').update(buffer).digest('hex');
try {
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const out = join(work, 'xlsx'),
    csv = join(work, 'csv');
  mkdirSync(out);
  mkdirSync(csv);
  const cases = [];
  for (const [name, rows] of Object.entries({
    blank: [['first'], [''], ['']],
    control: [['first'], [''], ['last']],
  })) {
    const input = join(work, `${name}.tsv`);
    writeFileSync(input, encodeClipboardText(rows));
    const xlsxLog = execFileSync(
      'soffice',
      ['--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', out, input],
      { encoding: 'utf8', timeout: 120000 },
    );
    const xlsx = join(out, `${name}.xlsx`);
    if (!existsSync(xlsx)) throw new Error(`Missing XLSX: ${xlsxLog}`);
    const csvLog = execFileSync(
      'soffice',
      ['--headless', '--convert-to', 'csv', '--outdir', csv, xlsx],
      { encoding: 'utf8', timeout: 120000 },
    );
    const csvPath = join(csv, `${name}.csv`);
    if (!existsSync(csvPath)) throw new Error(`Missing CSV: ${csvLog}`);
    const workbook = XLSX.read(readFileSync(xlsx), { type: 'buffer' }),
      sheet = workbook.Sheets[workbook.SheetNames[0]];
    const artifacts = [input, xlsx, csvPath].map((source) => {
      const target = `${prefix}-${name}.${source.split('.').at(-1)}`;
      mkdirSync(dirname(target), { recursive: true });
      copyFileSync(source, target);
      return { name: basename(target), sha256: hash(readFileSync(target)) };
    });
    cases.push({
      name,
      inputText: readFileSync(input, 'utf8'),
      range: sheet['!ref'] ?? null,
      cells: ['A1', 'A2', 'A3'].map((address) => ({
        address,
        type: sheet[address]?.t ?? null,
        value: sheet[address]?.v ?? null,
      })),
      csvText: readFileSync(csvPath, 'utf8'),
      xlsxLog,
      csvLog,
      artifacts,
    });
  }
  writeFileSync(
    output,
    JSON.stringify(
      {
        reader,
        node: process.version,
        source: sourceFingerprint(),
        encoderSha256: hash(bundled.outputFiles[0].text),
        cases,
        limitations:
          'TSV file import and XLSX/CSV save using LibreOfficeDev headless defaults; XLSX inspected using SheetJS. This does not exercise native clipboard paste into existing nonempty cells, Excel, or other locales.',
      },
      null,
      2,
    ) + '\n',
    { flag: 'wx' },
  );
  console.log(JSON.stringify({ reader, cases }));
} finally {
  rmSync(work, { recursive: true, force: true });
}
