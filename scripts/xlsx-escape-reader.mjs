import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const prefix = process.argv[2];
if (!prefix) throw new Error('Usage: node scripts/xlsx-escape-reader.mjs NEW_PREFIX');
const workbook = XLSX.utils.book_new();
const values = ['_x0041_', '_x005F_x0041_', '_x0000_', '_xZZZZ_', 'before\0after', 'before\rafter'];
const sheet = XLSX.utils.aoa_to_sheet([values]);
sheet.A1.c = [{ a: 'OpenSheet', t: 'note _x0041_' }];
sheet.A1.l = { Target: 'https://example.com/_x0041_' };
XLSX.utils.book_append_sheet(workbook, sheet, '_x0041_');
const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
const parsed = XLSX.read(bytes, { type: 'buffer' });
const temp = mkdtempSync(join(tmpdir(), 'opensheet-xlsx-escape-'));
try {
  const source = resolve(`${prefix}.xlsx`),
    csv = resolve(`${prefix}.csv`),
    reportPath = resolve(`${prefix}.json`);
  writeFileSync(source, bytes, { flag: 'wx' });
  const version = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${pathToFileURL(join(temp, 'profile')).href}`,
      '--headless',
      '--convert-to',
      'csv',
      '--outdir',
      temp,
      source,
    ],
    { encoding: 'utf8', timeout: 60000 },
  );
  const csvBytes = readFileSync(join(temp, basename(source).replace(/\.xlsx$/, '.csv')));
  writeFileSync(csv, csvBytes, { flag: 'wx' });
  const sha256 = (data) => createHash('sha256').update(data).digest('hex');
  const actual = parsed.Sheets[parsed.SheetNames[0]];
  const report = {
    reader: version,
    scope:
      'Synthetic SheetJS-generated inline string, note, hyperlink and worksheet name; LibreOffice headless CSV observes cell text only, while SheetJS reads the generated XLSX. Does not establish Excel behavior or independent note/link behavior.',
    artifacts: {
      xlsx: { file: basename(source), sha256: sha256(bytes) },
      csv: { file: basename(csv), sha256: sha256(csvBytes) },
    },
    inputs: values,
    libreOfficeCsv: csvBytes.toString('utf8'),
    sheetJS: {
      sheetNames: parsed.SheetNames,
      cells: ['A1', 'B1', 'C1', 'D1', 'E1', 'F1'].map((cell) => actual[cell].v),
      note: actual.A1.c[0].t,
      target: actual.A1.l.Target,
    },
  };
  writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify(report));
} finally {
  rmSync(temp, { recursive: true, force: true });
}
