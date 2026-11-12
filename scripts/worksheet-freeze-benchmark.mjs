import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/worksheet-freeze-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: [resolve('packages/adapter-sheetjs/src/index.ts')],
  bundle: true,
  platform: 'node',
  format: 'esm',
  write: false,
});
const { fromSheetJS } = await import(
  `data:text/javascript;base64,${Buffer.from(bundle.outputFiles[0].contents).toString('base64')}`
);
const cells = Array.from(
  { length: 100_000 },
  (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`,
).join('');
const rules = '<dataValidation sqref="A1"/><conditionalFormatting sqref="A1"/>'.repeat(1000);
const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const pkg = 'http://schemas.openxmlformats.org/package/2006/relationships';
const view =
  '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" state="frozen"/></sheetView></sheetViews>';
const xml = (frozen) =>
  `<worksheet xmlns="${main}">${frozen ? view : ''}<sheetData>${cells}</sheetData>${rules}</worksheet>`;
const paneContent = new TextEncoder().encode(xml(true));
const noPaneContent = new TextEncoder().encode(xml(false));
const file = (text) => ({ content: new TextEncoder().encode(text) });
const base = {
  SheetNames: ['Data'],
  Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
  Workbook: { Sheets: [{ name: 'Data', id: 'r1' }] },
  files: {
    'xl/workbook.xml': file(
      `<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Data" r:id="r1"/></sheets></workbook>`,
    ),
    'xl/_rels/workbook.xml.rels': file(
      `<Relationships xmlns="${pkg}"><Relationship Id="r1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
    ),
  },
};
const input = (content) => ({
  ...base,
  files: { ...base.files, 'xl/worksheets/sheet1.xml': { content } },
});
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  xmlBytes: paneContent.byteLength,
  xmlCells: 100_000,
  rulesPerType: 1000,
  limitations:
    'In-process Node conversion of a one-cell parsed worksheet; synthetic 100,000-cell raw XML with or without a frozen pane is only inspected, not parsed into cells. No ZIP, browser, Worker, full import or isolated peak-memory measurement.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 5; round++) {
  for (const variant of round % 2 ? ['pane', 'no-pane'] : ['no-pane', 'pane']) {
    const start = performance.now();
    const result = fromSheetJS(input(variant === 'pane' ? paneContent : noPaneContent));
    const elapsedMs = performance.now() - start;
    for (const feature of ['data-validation', 'conditional-formatting']) {
      const issue = result.report.issues.find((item) => item.feature === feature);
      if (issue?.count !== 1000) throw new Error(`Unexpected ${feature} report in ${variant}`);
    }
    const freeze = result.file.sheets[0].freeze;
    if (
      freeze.rows !== (variant === 'pane' ? 3 : 0) ||
      freeze.columns !== (variant === 'pane' ? 2 : 0)
    )
      throw new Error('Unexpected freeze result');
    report.samples.push({ round: round + 1, variant, elapsedMs });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report.samples.at(-1)));
  }
}
