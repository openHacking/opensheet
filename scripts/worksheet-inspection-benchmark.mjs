import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/worksheet-inspection-benchmark.mjs NEW_REPORT');
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
const xml = `<worksheet><sheetData>${cells}</sheetData>${rules}<!--<dataValidation/>--></worksheet>`;
const content = new TextEncoder().encode(xml);
const base = {
  SheetNames: ['Data'],
  Sheets: { Data: { A1: { t: 'n', v: 1 }, '!ref': 'A1' } },
};
const input = {
  ...base,
  files: { 'xl/worksheets/sheet1.xml': { content } },
};
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  xmlBytes: content.byteLength,
  xmlCells: 100_000,
  rulesPerType: 1000,
  limitations:
    'In-process Node conversion of a one-cell parsed worksheet; synthetic 100,000-cell raw XML is only inspected, not parsed into cells. No ZIP, browser, Worker, full import or isolated peak-memory measurement.',
  samples: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
for (let round = 0; round < 5; round++) {
  for (const variant of round % 2 ? ['inspection', 'no-xml'] : ['no-xml', 'inspection']) {
    const start = performance.now();
    const result = fromSheetJS(variant === 'inspection' ? input : base);
    const elapsedMs = performance.now() - start;
    for (const feature of ['data-validation', 'conditional-formatting']) {
      const issue = result.report.issues.find((item) => item.feature === feature);
      if (variant === 'inspection' ? issue?.count !== 1000 : issue !== undefined)
        throw new Error(`Unexpected ${feature} report in ${variant}`);
    }
    report.samples.push({ round: round + 1, variant, elapsedMs });
    writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report.samples.at(-1)));
  }
}
