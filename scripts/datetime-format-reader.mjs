import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  mkdtempSync,
  mkdirSync,
  readFileSync,
  writeFileSync,
  existsSync,
  rmSync,
  copyFileSync,
  constants,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { createWorkbookFile, encodeFileCells, formatValue } from '../packages/core/dist/index.js';
import { toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { parseDelimited } from '../packages/formats/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';
const output = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!output) throw new Error('Supply a new --output path');
const reportPath = resolve(output),
  prefix = reportPath.replace(/\.json$/, '');
const artifacts = Object.fromEntries(
  ['1900', '1904'].flatMap((system) => [
    [`${system}xlsx`, `${prefix}.${system}.xlsx`],
    [`${system}csv`, `${prefix}.${system}.csv`],
  ]),
);
if ([reportPath, ...Object.values(artifacts)].some(existsSync))
  throw new Error('Refusing to overwrite evidence');
const work = mkdtempSync(join(tmpdir(), 'opensheet-datetime-'));
const reader = execFileSync('soffice', ['--version'], { encoding: 'utf8', timeout: 30000 }).trim();
const cases = [
  { value: 45293.5, format: 'yyyy-mm-dd hh:mm' },
  { value: 45293.999999, format: 'yyyy-mm-dd hh:mm' },
  { value: 0.5, format: 'hh:mm' },
  { value: 60.5, format: 'yyyy-mm-dd hh:mm:ss' },
  { value: 0, format: 'm/d/yy' },
  { value: 60, format: 'mm/dd/yyyy' },
  { value: 45293.5, format: 'm/d/yy h:mm' },
  { value: 45293.999999, format: 'm/d/yy h:mm' },
];
try {
  const results = [];
  for (const system of ['1900', '1904']) {
    const dir = join(work, system);
    mkdirSync(dir);
    const file = createWorkbookFile({
      sheets: [{ name: 'Dates', rows: cases.length, columns: 1 }],
    });
    file.dateSystem = system;
    file.sheets[0].blocks = [
      {
        row: 0,
        column: 0,
        data: encodeFileCells(
          cases.map((item, row) => ({
            rowId: row,
            columnId: 0,
            input: { type: 'number', value: item.value },
            numberFormat: item.format,
          })),
          new Map(),
        ),
      },
    ];
    const source = join(dir, 'source.xlsx');
    writeFileSync(
      source,
      XLSX.write(toSheetJS(file).workbook, { type: 'buffer', bookType: 'xlsx' }),
    );
    const log = execFileSync(
      'soffice',
      [
        `-env:UserInstallation=${pathToFileURL(join(work, 'profile')).href}`,
        '--headless',
        '--convert-to',
        'csv:Text - txt - csv (StarCalc):44,34,76,1',
        '--outdir',
        dir,
        source,
      ],
      { encoding: 'utf8', timeout: 120000 },
    ).trim();
    const csvPath = join(dir, 'source.csv'),
      rows = parseDelimited(readFileSync(csvPath, 'utf8'));
    const comparisons = cases.map((item, i) => ({
      ...item,
      opensheet: formatValue(item.value, item.format, system),
      reader: rows[i]?.[0] ?? null,
      agrees: rows[i]?.[0] === formatValue(item.value, item.format, system),
    }));
    copyFileSync(source, artifacts[`${system}xlsx`], constants.COPYFILE_EXCL);
    copyFileSync(csvPath, artifacts[`${system}csv`], constants.COPYFILE_EXCL);
    results.push({ system, comparisons, log });
  }
  const report = {
    date: new Date().toISOString(),
    reader,
    provenance:
      'CC0 synthetic numeric date/time cells exported through OpenSheet and SheetJS; display text independently produced by LibreOffice CSV conversion. Excel and visual layout not tested.',
    source: sourceFingerprint(),
    build: localBuildFingerprint(),
    results,
    artifacts: Object.fromEntries(
      Object.entries(artifacts).map(([key, path]) => [
        key,
        {
          file: basename(path),
          sha256: createHash('sha256').update(readFileSync(path)).digest('hex'),
        },
      ]),
    ),
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  process.stdout.write(JSON.stringify({ reader, results }, null, 2) + '\n');
} finally {
  rmSync(work, { recursive: true, force: true });
}
