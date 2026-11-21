import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, relative, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import * as XLSX from 'xlsx';
import { fromSheetJS, toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { FileReader } from '../packages/core/dist/index.js';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const [output, directory] = process.argv.slice(2);
if (!output?.endsWith('.json') || !directory || process.argv.length !== 4)
  throw Error('Usage: node scripts/date-system-reader.mjs NEW_REPORT.json NEW_ARTIFACT_DIRECTORY');
const root = process.cwd(),
  artifacts = resolve(directory),
  temp = mkdtempSync(join(tmpdir(), 'opensheet-date-reader-'));
const hash = (data) => createHash('sha256').update(data).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  node: process.version,
  scope:
    'Two synthetic openpyxl workbooks; OpenSheet adapter export; independent LibreOfficeDev headless save/recalculation and openpyxl inspection. No Excel, stable LibreOffice, visual review or full date compatibility.',
  systems: [],
};
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
try {
  mkdirSync(artifacts, { recursive: false });
  const readerDir = join(artifacts, 'reader');
  mkdirSync(readerDir);
  report.build = Object.fromEntries(
    ['core', 'formula', 'adapter-sheetjs'].map((name) => [
      name,
      hash(readFileSync(join(root, 'packages', name, 'dist/index.js'))),
    ]),
  );
  report.inspectorSha256 = hash(readFileSync(join(root, 'scripts/xlsx-date-system-inspect.py')));
  report.readerVersion = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const profile = pathToFileURL(join(temp, 'lo-profile')).href;
  const inspect = (paths) =>
    JSON.parse(
      execFileSync('python3', [join(root, 'scripts/xlsx-date-system-inspect.py'), ...paths], {
        encoding: 'utf8',
        timeout: 30000,
      }),
    );
  const parse = (bytes) =>
    XLSX.read(bytes, {
      type: 'buffer',
      cellNF: true,
      cellStyles: true,
      sheetStubs: true,
      bookFiles: true,
    });
  const observe = (bytes) => {
    const result = fromSheetJS(parse(bytes)),
      reader = new FileReader(result.file);
    try {
      const sheet = reader.getSheetById(result.file.sheets[0].id);
      return {
        system: result.file.dateSystem,
        values: sheet.range('A1:A3').getValues(),
        display: sheet.range('A1:A3').getDisplayValues(),
        formula: reader.getCell(sheet.id, 0, 1)?.input,
        calculated: sheet.range('B1').getValues()[0][0],
        report: result.report,
      };
    } finally {
      reader.dispose();
    }
  };
  for (const system of ['1900', '1904']) {
    const source = join(root, 'benchmarks/date-system-openpyxl-2026-11-22', `dates-${system}.xlsx`);
    const evidence = JSON.parse(readFileSync(`${source}.json`, 'utf8'));
    const sourceBytes = readFileSync(source);
    assert.equal(hash(sourceBytes), evidence.sha256);
    const imported = fromSheetJS(parse(sourceBytes));
    const exported = toSheetJS(imported.file);
    const native = join(artifacts, `opensheet-${system}.xlsx`);
    writeFileSync(native, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }), {
      flag: 'wx',
    });
    const command = execFileSync(
      'soffice',
      [
        '--headless',
        `-env:UserInstallation=${profile}`,
        '--convert-to',
        'xlsx',
        '--outdir',
        readerDir,
        source,
        native,
      ],
      { encoding: 'utf8', timeout: 120000, stdio: ['ignore', 'pipe', 'pipe'] },
    );
    const control = join(readerDir, `dates-${system}.xlsx`),
      roundtrip = join(readerDir, `opensheet-${system}.xlsx`);
    const paths = [source, native, control, roundtrip];
    const independent = inspect(paths);
    const openSheet = paths.map((path) => observe(readFileSync(path)));
    assert.equal(independent.files[0].cached, null, 'Source formula unexpectedly had a cache');
    assert.equal(
      independent.files[1].cached,
      null,
      'OpenSheet export unexpectedly had a formula cache',
    );
    const expectedDates = ['1904-01-02T00:00:00', '2024-02-29T12:30:00', '2100-03-01T00:00:00'];
    const expectedDisplay = [['1904-01-02'], ['2024-02-29 12:30'], ['2100-03-01']];
    for (let index = 0; index < paths.length; index++) {
      const observed = independent.files[index];
      assert.equal(observed.system, system);
      assert.deepEqual(observed.dates, expectedDates);
      assert.equal(observed.formula, '=A2-A1');
      if (index >= 2)
        assert.ok(
          Math.abs(observed.cached - (evidence.xlsxSerials[1] - evidence.xlsxSerials[0])) < 1e-8,
        );
      assert.equal(openSheet[index].system, system);
      assert.deepEqual(openSheet[index].display, expectedDisplay);
      assert.ok(
        Math.abs(
          openSheet[index].calculated - (evidence.xlsxSerials[1] - evidence.xlsxSerials[0]),
        ) < 1e-8,
      );
    }
    report.systems.push({
      system,
      command,
      inputSha256: evidence.sha256,
      files: paths.map((path) => ({
        path: relative(root, path),
        sha256: hash(readFileSync(path)),
      })),
      independent,
      openSheet,
      exportReport: exported.report,
    });
    save();
  }
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  if (error.stdout) report.stdout = String(error.stdout).slice(-5000);
  if (error.stderr) report.stderr = String(error.stderr).slice(-5000);
  save();
  throw error;
} finally {
  rmSync(temp, { recursive: true, force: true });
}
