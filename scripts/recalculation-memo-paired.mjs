import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/recalculation-memo-paired.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: [resolve('packages/core/src/async-calculation.ts')],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  globalName: 'CalculationBenchmark',
  write: false,
});
const workerSource =
  bundle.outputFiles[0].text +
  `
self.onmessage = async ({data}) => {
  try {
    let sourceReads = 0;
    const calc = new CalculationBenchmark.AsyncCalculation(async (sheet, row, column) => {
      if (sheet === 'source') { sourceReads++; return { rowId: row, columnId: column, input: { type:'number', value:1 } }; }
      return { rowId:row, columnId:column, input:{ type:'formula', expression:'SUM(Source!$A$1:$A$32)+' + row } };
    }, (name) => name === 'Source' ? 'source' : undefined);
    const read = data.memo ? calc.readSession() : (sheet, row, column) => calc.value(sheet, row, column);
    const start = performance.now();
    for (let row = 0; row < data.rows; row++) {
      const value = await read('formula', row, 0);
      if (value !== 32 + row) throw new Error('Wrong output ' + row);
    }
    const ms = performance.now() - start;
    const expectedReads = data.memo ? 32 : data.rows * 32;
    if (sourceReads !== expectedReads) throw new Error('Wrong source read count ' + sourceReads);
    postMessage({ memo:data.memo, rows:data.rows, ms, sourceReads, checkedOutputs:data.rows });
  } catch (error) { postMessage({error:String(error)}); }
};`;
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hash: createHash('sha256').update(workerSource).digest('hex'),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless dedicated Worker synthetic AsyncCalculation benchmark. Six alternating pairs per browser and 128/512 formulas over the same 32 scalar references; fresh calculator each sample, all outputs and exact source-read counts verified. Baseline uses the current calculator without a read session; this is not a prior commit. Excludes IndexedDB, real workbook scheduling, UI, cancellation, peak memory, reference devices and end-to-end latency.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    for (let round = 0; round < 6; round++) {
      result.rounds.push(
        await page.evaluate(
          async ({ workerSource, round }) => {
            const url = URL.createObjectURL(new Blob([workerSource], { type: 'text/javascript' }));
            const worker = new Worker(url);
            try {
              const samples = [];
              for (const memo of round % 2 ? [true, false] : [false, true]) {
                for (const rows of round % 2 ? [512, 128] : [128, 512]) {
                  const sample = await new Promise((resolve, reject) => {
                    worker.onmessage = ({ data }) =>
                      data.error ? reject(new Error(data.error)) : resolve(data);
                    worker.onerror = (error) => reject(new Error(error.message));
                    worker.postMessage({ memo, rows });
                  });
                  samples.push(sample);
                }
              }
              return samples;
            } finally {
              worker.terminate();
              URL.revokeObjectURL(url);
            }
          },
          { workerSource, round },
        ),
      );
      save();
    }
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
