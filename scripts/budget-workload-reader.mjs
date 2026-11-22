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
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { transformSync } from 'esbuild';
import * as XLSX from 'xlsx';
import { toSheetJS } from '../packages/adapter-sheetjs/dist/index.js';
import { localBuildFingerprint, sourceFingerprint } from './benchmark-fingerprint.mjs';

const arg = process.argv.find((value) => value.startsWith('--output='))?.slice(9);
if (!arg?.endsWith('.json')) throw new Error('Pass --output=NEW_REPORT.json');
const output = resolve(arg),
  prefix = output.slice(0, -5);
const targets = [
  'initial.source.xlsx',
  'initial.saved.xlsx',
  'edited.source.xlsx',
  'edited.saved.xlsx',
].map((suffix) => `${prefix}.${suffix}`);
if ([output, ...targets].some(existsSync)) throw new Error('Refusing to overwrite evidence');
const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const fixturePath = 'apps/playground/src/budget-fixture.ts';
const fixtureBytes = readFileSync(fixturePath);
const code = transformSync(fixtureBytes.toString('utf8'), { loader: 'ts', format: 'esm' }).code;
const fixture = (
  await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`)
).default;
const reader = execFileSync('soffice', ['--version'], { encoding: 'utf8', timeout: 30000 }).trim();
const work = mkdtempSync(join(tmpdir(), 'opensheet-budget-reader-'));
try {
  const observations = [];
  for (const [index, quantity] of [1, 3].entries()) {
    const file = structuredClone(fixture);
    const [budget, notes] = file.sheets;
    const numbers = budget.blocks[0].data.numbers;
    const position = numbers.positions.indexOf(35);
    if (position < 0 || numbers.values[position] !== 1)
      throw new Error('Unexpected source quantity');
    numbers.values[position] = quantity;
    const data = notes.blocks[0].data;
    if (data.formulas) throw new Error('Unexpected existing notes formulas');
    data.formulas = {
      positions: [35, 36],
      values: [
        "'Launch budget'!$D2*'Launch budget'!E$2",
        "'Launch budget'!$D$2+'Launch budget'!E$2",
      ],
    };
    const exported = toSheetJS(file);
    const input = join(work, `case-${index}.xlsx`),
      dest = join(work, `out-${index}`);
    mkdirSync(dest);
    writeFileSync(input, XLSX.write(exported.workbook, { type: 'buffer', bookType: 'xlsx' }));
    const source = XLSX.read(readFileSync(input), { type: 'buffer', cellFormula: true });
    const cells = [
      { sheet: 'Launch budget', address: 'D2', expected: quantity },
      { sheet: 'Launch budget', address: 'F2', expected: quantity * 2400 },
      { sheet: 'Launch budget', address: 'F11', expected: quantity === 1 ? 13097 : 17897 },
      { sheet: 'Notes', address: 'D2', expected: quantity * 2400 },
      { sheet: 'Notes', address: 'E2', expected: quantity + 2400 },
    ];
    const log = execFileSync(
      'soffice',
      [
        `-env:UserInstallation=${pathToFileURL(join(work, `profile-${index}`)).href}`,
        '--headless',
        '--convert-to',
        'xlsx:Calc MS Excel 2007 XML',
        '--outdir',
        dest,
        input,
      ],
      { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], timeout: 120000 },
    );
    const saved = join(dest, basename(input));
    if (!existsSync(saved)) throw new Error(`Reader produced no file: ${log}`);
    const result = XLSX.read(readFileSync(saved), { type: 'buffer', cellFormula: true });
    const values = cells.map((cell) => {
      const original = source.Sheets[cell.sheet]?.[cell.address];
      const actual = result.Sheets[cell.sheet]?.[cell.address];
      return {
        ...cell,
        sourceFormula: original?.f ?? null,
        readerFormula: actual?.f ?? null,
        readerValue: actual?.v ?? null,
        readerType: actual?.t ?? null,
        agrees: actual?.t !== 'e' && actual?.v === cell.expected,
      };
    });
    copyFileSync(input, targets[index * 2]);
    copyFileSync(saved, targets[index * 2 + 1]);
    observations.push({ quantity, values, report: exported.report, log });
  }
  const report = {
    reader,
    node: process.version,
    source: sourceFingerprint(),
    adapterBuild: localBuildFingerprint('packages/adapter-sheetjs/dist'),
    fixture: { file: fixturePath, sha256: hash(fixtureBytes) },
    observations,
    artifacts: targets.map((path) => ({ name: basename(path), sha256: hash(readFileSync(path)) })),
    limitations:
      'Two synthetic variants of the checked-in budget fixture; SheetJS writes and inspects XLSX, while one LibreOfficeDev alpha build recalculates it headlessly. Does not cover native editing, Excel, stable LibreOffice, locales or other formulas.',
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(
    JSON.stringify(
      { reader, values: observations.map(({ quantity, values }) => ({ quantity, values })) },
      null,
      2,
    ),
  );
} finally {
  rmSync(work, { recursive: true, force: true });
}
