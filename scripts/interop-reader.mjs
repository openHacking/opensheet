import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  readFileSync,
  writeFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  copyFileSync,
  constants,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve, join, dirname, basename } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const outputArg = process.argv.find((arg) => arg.startsWith('--output='))?.slice(9);
if (!outputArg)
  throw new Error('Pass --output=new-report.json; historical evidence is never overwritten');
const output = resolve(outputArg),
  prefix = output.replace(/\.json$/, '');
const artifacts = {
  source: prefix + '.source.fods',
  authored: prefix + '.authored.xlsx',
  exported: prefix + '.opensheet.xlsx',
  reopened: prefix + '.reader.xlsx',
};
if ([output, ...Object.values(artifacts)].some(existsSync))
  throw new Error('Evidence already exists');
const manifestPath = resolve('tests/fixtures/interop/manifest.json');
const fixture = JSON.parse(readFileSync(manifestPath, 'utf8')).fixtures.find(
  (entry) => entry.id === 'basic',
);
if (!fixture) throw new Error('Missing basic fixture manifest');
const source = resolve(dirname(manifestPath), fixture.source);
const fingerprints = {
  source: sourceFingerprint(),
  adapterBuild: localBuildFingerprint('packages/adapter-sheetjs/dist'),
  manifestSha256: createHash('sha256').update(readFileSync(manifestPath)).digest('hex'),
};
const work = mkdtempSync(join(tmpdir(), 'opensheet-interop-'));
const digest = (bytes) => createHash('sha256').update(bytes).digest('hex');
const expected = fixture.expectedValues;
function inspect(bytes) {
  const book = XLSX.read(bytes, {
    type: 'buffer',
    cellStyles: true,
    cellNF: true,
    cellDates: false,
    xlfn: true,
  });
  return {
    book,
    observation: {
      sheetNames: book.SheetNames,
      cells: Object.fromEntries(
        Object.keys(expected).map((key) => {
          const [sheet, address] = key.split('!'),
            cell = book.Sheets[sheet]?.[address];
          return [
            key,
            {
              value: cell?.v ?? null,
              type: cell?.t,
              formula: cell?.f,
              format: cell?.z,
              note: cell?.c?.map((c) => c.t).join('\n'),
              link: cell?.l?.Target,
              style: cell?.s,
            },
          ];
        }),
      ),
      merges: book.Sheets.Data?.['!merges'] ?? [],
      rows: book.Sheets.Data?.['!rows'] ?? [],
      columns: book.Sheets.Data?.['!cols'] ?? [],
    },
  };
}
try {
  const reader = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const logs = [];
  function convert(input, folder) {
    const destination = join(work, folder);
    mkdirSync(destination);
    logs.push(
      execFileSync(
        'soffice',
        [
          `-env:UserInstallation=${pathToFileURL(join(work, 'profile')).href}`,
          '--headless',
          '--convert-to',
          'xlsx:Calc MS Excel 2007 XML',
          '--outdir',
          destination,
          input,
        ],
        { encoding: 'utf8', timeout: 120000 },
      ).trim(),
    );
    const result = join(destination, basename(input).replace(/\.[^.]+$/, '.xlsx'));
    if (!existsSync(result)) throw new Error(`No reader output: ${logs.at(-1)}`);
    return result;
  }
  const authored = convert(source, 'authored');
  const imported = inspect(readFileSync(authored));
  const native = fromSheetJS(imported.book),
    exported = toSheetJS(native.file);
  const exportedPath = join(work, 'opensheet.xlsx');
  writeFileSync(exportedPath, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
  const reopened = convert(exportedPath, 'reopened');
  const final = inspect(readFileSync(reopened));
  mkdirSync(dirname(output), { recursive: true });
  copyFileSync(authored, artifacts.authored, constants.COPYFILE_EXCL);
  copyFileSync(source, artifacts.source, constants.COPYFILE_EXCL);
  copyFileSync(exportedPath, artifacts.exported, constants.COPYFILE_EXCL);
  copyFileSync(reopened, artifacts.reopened, constants.COPYFILE_EXCL);
  const assertions = Object.entries(expected).map(([cell, value]) => ({
    cell,
    expected: value,
    authored: imported.observation.cells[cell].value,
    reopened: final.observation.cells[cell].value,
    agrees:
      imported.observation.cells[cell].value === value &&
      final.observation.cells[cell].value === value,
  }));
  const report = {
    schemaVersion: 1,
    date: new Date().toISOString(),
    reader,
    fingerprints,
    provenance:
      'CC0 synthetic hand-authored Flat ODF source, independently converted to XLSX by LibreOffice; imported/exported through OpenSheet adapter; reopened and saved by LibreOffice. Expected cell values are hand-authored, not generated by OpenSheet. SheetJS decodes observations; visual fidelity is not established.',
    source: 'tests/fixtures/interop/basic.fods',
    sourceSha256: digest(readFileSync(source)),
    artifacts: Object.fromEntries(
      Object.entries(artifacts).map(([name, path]) => [
        name,
        { file: basename(path), sha256: digest(readFileSync(path)) },
      ]),
    ),
    assertions,
    authored: imported.observation,
    reopened: final.observation,
    importReport: native.report,
    exportReport: exported.report,
    logs,
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  process.stdout.write(
    JSON.stringify(
      {
        output,
        reader,
        agreements: assertions.filter((a) => a.agrees).length,
        assertions: assertions.length,
        differences: assertions.filter((a) => !a.agrees),
        importSummary: native.report.summary,
        exportSummary: exported.report.summary,
      },
      null,
      2,
    ) + '\n',
  );
} finally {
  rmSync(work, { recursive: true, force: true });
}
