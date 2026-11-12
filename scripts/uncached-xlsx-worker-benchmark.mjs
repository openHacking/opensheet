import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/uncached-xlsx-worker-benchmark.mjs NEW_REPORT');
const bytes = readFileSync('benchmarks/uncached-import-10k-2026-11-11.xlsx');
const result = await build({
  stdin: {
    contents: `
import * as XLSX from 'xlsx';
self.onmessage = ({data}) => {
  const start = performance.now();
  const workbook = XLSX.read(data.bytes, {type: 'array', dense: true, cellFormula: true,
    cellNF: true, cellStyles: data.styles, sheetStubs: data.stubs});
  const ms = performance.now() - start;
  const rows = workbook.Sheets['Uncached benchmark']['!data'];
  let numbers = 0, formulas = 0, formattedBlanks = 0;
  for (let i = 0; i < 10000; i++) {
    const row = rows[i];
    if (row?.[0]?.v === i + 1) numbers++;
    if (row?.[1]?.f === 'A' + (i + 1) + '*2') formulas++;
    if (i % 5 === 4 && row?.[2]?.z === '0.00') formattedBlanks++;
  }
  if (numbers !== 10000 || formulas !== (data.styles || data.stubs ? 10000 : 0) ||
      formattedBlanks !== (data.styles || data.stubs ? 2000 : 0))
    throw new Error('Unexpected data: ' + JSON.stringify({numbers, formulas, formattedBlanks}));
  postMessage({ms, numbers, formulas, formattedBlanks});
};`,
    resolveDir: process.cwd(),
    loader: 'js',
  },
  bundle: true,
  platform: 'browser',
  format: 'iife',
  write: false,
});
const script = result.outputFiles[0].text;
const hash = (value) => createHash('sha256').update(value).digest('hex');
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  workerSha256: hash(script),
  inputSha256: hash(bytes),
  inputBytes: bytes.length,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless dedicated Worker SheetJS parse only, three alternating fresh-Worker rounds per browser. An openpyxl fixture contains 10,000 uncached formulas, 10,000 numeric cells and 2,000 formatted blanks. All cells checked outside timing. Styles-on profiles retain cells with either stubs setting; styles-off/stubs-off is an incomplete-data comparison. Excludes ZIP preflight, adapter, persistence, UI, peak memory, devices and other workbooks.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const entry = { engine, version: browser.version(), rounds: [] };
  report.results.push(entry);
  try {
    const page = await browser.newPage();
    for (let round = 0; round < 3; round++) {
      const current = { round: round + 1, samples: [] };
      entry.rounds.push(current);
      try {
        current.samples = await page.evaluate(
          async ({ script, bytes, round }) => {
            const profiles = [
              { name: 'styles-no-stubs', styles: true, stubs: false },
              { name: 'styles-stubs', styles: true, stubs: true },
              { name: 'no-styles-no-stubs', styles: false, stubs: false },
            ];
            if (round % 2) profiles.reverse();
            const samples = [];
            for (const profile of profiles) {
              const url = URL.createObjectURL(new Blob([script], { type: 'text/javascript' }));
              const worker = new Worker(url);
              try {
                const value = await new Promise((resolve, reject) => {
                  worker.onmessage = ({ data }) => resolve(data);
                  worker.onerror = reject;
                  worker.postMessage({ ...profile, bytes: Uint8Array.from(bytes).buffer });
                });
                samples.push({ profile: profile.name, ...value });
              } finally {
                worker.terminate();
                URL.revokeObjectURL(url);
              }
            }
            return samples;
          },
          { script, bytes: Array.from(bytes), round },
        );
        console.log(JSON.stringify({ engine, ...current }));
      } catch (error) {
        current.error = String(error);
        throw error;
      } finally {
        save();
      }
    }
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
