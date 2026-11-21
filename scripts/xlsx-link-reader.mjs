import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { createWorkbookFile, encodeFileCells } from '../packages/core/dist/index.js';
import { toSheetJS, writeXlsx } from '../packages/adapter-sheetjs/dist/index.js';

const prefix = process.argv[2];
if (!prefix) throw new Error('Usage: node scripts/xlsx-link-reader.mjs NEW_PREFIX');
const paths = {
  baseline: resolve(`${prefix}.baseline.xlsx`),
  candidate: resolve(`${prefix}.xlsx`),
  saved: resolve(`${prefix}.saved.xlsx`),
  report: resolve(`${prefix}.json`),
};
const targets = [
  'https://例子.测试/é/🙂?x=1&y=2#片段',
  'mailto:用户@example.com?subject=你好',
  'https://example.com/a%20b?x=1&y=2#frag',
  'http://EXAMPLE.com:80/a/../b',
];
const file = createWorkbookFile({ sheets: [{ name: 'Links', rows: targets.length, columns: 1 }] });
file.sheets[0].blocks = [
  {
    row: 0,
    column: 0,
    data: encodeFileCells(
      targets.map((target, rowId) => ({
        rowId,
        columnId: 0,
        input: { type: 'string', value: 'link' },
        link: { target },
      })),
      new Map(),
    ),
  },
];
const candidate = writeXlsx(file, XLSX);
const baselineBook = toSheetJS(file).workbook;
for (let i = 0; i < targets.length; i++)
  baselineBook.Sheets.Links[`A${i + 1}`].l.Target = targets[i];
const baseline = XLSX.write(baselineBook, { type: 'buffer', bookType: 'xlsx' });
const temp = mkdtempSync(join(tmpdir(), 'opensheet-xlsx-link-'));
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
try {
  writeFileSync(paths.baseline, baseline, { flag: 'wx' });
  writeFileSync(paths.candidate, candidate.bytes, { flag: 'wx' });
  const version = execFileSync('soffice', ['--version'], {
    encoding: 'utf8',
    timeout: 30000,
  }).trim();
  const log = execFileSync(
    'soffice',
    [
      `-env:UserInstallation=${pathToFileURL(join(temp, 'profile')).href}`,
      '--headless',
      '--convert-to',
      'xlsx',
      '--outdir',
      temp,
      paths.candidate,
    ],
    { encoding: 'utf8', timeout: 60000 },
  ).trim();
  const saved = readFileSync(join(temp, basename(paths.candidate)));
  writeFileSync(paths.saved, saved, { flag: 'wx' });
  const python = `import json,openpyxl,sys
book=openpyxl.load_workbook(sys.argv[1]); sheet=book['Links']
print(json.dumps({'version':openpyxl.__version__,'links':[{'target':sheet[f'A{i}'].hyperlink.target,'location':sheet[f'A{i}'].hyperlink.location} for i in range(1,5)]},ensure_ascii=True))`;
  const inspect = (path) =>
    JSON.parse(execFileSync('python3', ['-c', python, path], { encoding: 'utf8', timeout: 30000 }));
  const sheetjs = (bytes) => {
    const sh = XLSX.read(bytes, { type: 'array' }).Sheets.Links;
    return targets.map((_, i) => sh[`A${i + 1}`].l.Target);
  };
  const oldTargets = sheetjs(baseline);
  const newTargets = sheetjs(candidate.bytes);
  if (oldTargets[0] === targets[0] || oldTargets[1] === targets[1])
    throw new Error('Baseline did not reproduce the Unicode relationship decoding issue');
  if (
    newTargets[0] !== new URL(targets[0]).href ||
    newTargets[1] !== new URL(targets[1]).href ||
    newTargets[2] !== targets[2] ||
    newTargets[3] !== targets[3]
  )
    throw new Error('Candidate targets differ from expected');
  const before = inspect(paths.candidate);
  const after = inspect(paths.saved);
  const report = {
    reader: version,
    scope:
      'Synthetic OpenSheet XLSX; baseline reconstructs prior raw link writing from the current SheetJS workbook. SheetJS reads both source files; openpyxl inspects candidate before and after LibreOfficeDev save. This does not establish Excel, stable LibreOffice, actual network navigation or all URI semantics.',
    inputs: targets,
    adapterReport: candidate.report,
    sheetjs: { baseline: oldTargets, candidate: newTargets },
    sourceInspection: before,
    savedInspection: after,
    artifacts: {
      baseline: { file: basename(paths.baseline), sha256: sha256(baseline) },
      candidate: { file: basename(paths.candidate), sha256: sha256(candidate.bytes) },
      saved: { file: basename(paths.saved), sha256: sha256(saved) },
    },
    versions: { node: process.version, openpyxl: before.version },
    source: {
      commit: execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim(),
      adapterBundleSha256: sha256(readFileSync('packages/adapter-sheetjs/dist/index.js')),
      adapterCellsSourceSha256: sha256(readFileSync('packages/adapter-sheetjs/src/cells.ts')),
      scriptSha256: sha256(readFileSync('scripts/xlsx-link-reader.mjs')),
      lockfileSha256: sha256(readFileSync('pnpm-lock.yaml')),
    },
    conversionLog: log,
  };
  writeFileSync(paths.report, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify({ baseline: oldTargets, candidate: newTargets, before, after }));
} finally {
  rmSync(temp, { recursive: true, force: true });
}
