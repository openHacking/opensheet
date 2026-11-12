import { build } from 'esbuild';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { performance } from 'node:perf_hooks';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Supply a new report path');
const file = resolve('packages/adapter-sheetjs/src/worksheet-losses.ts');
const source = readFileSync(file, 'utf8');
const marker =
  "    const inspectTree =\n      xml.includes('pane') || xml.includes('tabSelected') || xml.includes('<selection');";
if (source.split(marker).length !== 2) throw new Error('Cannot reconstruct previous view scan');
const variants = [source.replace(marker, "    const inspectTree = xml.includes('pane');"), source];
const modules = [],
  hashes = [];
for (const contents of variants) {
  const bundle = await build({
    entryPoints: [file],
    bundle: true,
    platform: 'node',
    format: 'esm',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /worksheet-losses\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/adapter-sheetjs/src'),
          }));
        },
      },
    ],
  });
  const code = bundle.outputFiles[0].text;
  hashes.push(createHash('sha256').update(code).digest('hex'));
  modules.push(await import(`data:text/javascript;base64,${Buffer.from(code).toString('base64')}`));
}
const cells = Array.from(
  { length: 100000 },
  (_, i) => `<row r="${i + 1}"><c r="A${i + 1}" t="n"><v>${i}</v></c></row>`,
).join('');
const root = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const fixtures = Object.fromEntries(
  Object.entries({
    active:
      '<sheetViews><sheetView workbookViewId="0" tabSelected="true"><selection activeCell="D8" sqref="D8"/></sheetView></sheetViews>',
    control: '<sheetViews><sheetView workbookViewId="0"/></sheetViews>',
  }).map(([name, view]) => [
    name,
    {
      'xl/worksheets/sheet1.xml': {
        content: new TextEncoder().encode(
          `<worksheet xmlns="${root}">${view}<sheetData>${cells}</sheetData></worksheet>`,
        ),
      },
    },
  ]),
);
const report = {
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  hashes,
  cells: 100000,
  scope:
    'Local Node paired XML-only scan, 100k synthetic scalar XML rows, six alternating pairs after warmup. Baseline reconstructs the pre-change inspection condition, not a clean commit. Does not parse workbook data, read ZIP, run in browsers/Workers or measure peak memory.',
  results: [],
};
for (const [name, files] of Object.entries(fixtures)) {
  const run = (index) => {
    const summary = { exact: 0, approximated: 0, dropped: 0, blocked: 0 },
      result = {
        adapterVersion: 'test',
        sourceFormat: 'SheetJS',
        summary,
        issues: [],
        omittedIssueOccurrences: 0,
      };
    const start = performance.now();
    modules[index].reportWorksheetLosses(files, result);
    const ms = performance.now() - start;
    const found = result.issues.map((issue) => issue.code).sort();
    const expected =
      name === 'active' && index === 1
        ? ['WORKSHEET_ACTIVE_TAB_DROPPED', 'WORKSHEET_SELECTION_DROPPED']
        : [];
    if (JSON.stringify(found) !== JSON.stringify(expected))
      throw new Error('Unexpected view report');
    return ms;
  };
  run(0);
  run(1);
  const pairs = [];
  for (let i = 0; i < 6; i++) {
    const pair = [];
    for (const variant of i % 2 ? [1, 0] : [0, 1]) pair[variant] = run(variant);
    pairs.push({ previous: pair[0], current: pair[1] });
  }
  report.results.push({ name, bytes: files['xl/worksheets/sheet1.xml'].content.length, pairs });
}
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
