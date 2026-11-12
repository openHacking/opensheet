import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/recalculation-memo-lookup-paired.mjs NEW_REPORT');
const path = resolve('packages/core/src/async-calculation.ts');
const previous = readFileSync(path, 'utf8');
const cached =
  '    const cached = this.cache.get(id);\n    if (cached !== undefined) return cached;\n';
const marker = "    if (stack.has(id)) return { error: '#CYCLE!' };\n";
if (previous.split(cached).length !== 2 || previous.split(marker).length !== 2)
  throw new Error('Cannot reconstruct prior lookup order');
const current = previous.replace(cached, '').replace(marker, cached + marker);
const sources = {},
  hashes = {};
for (const [variant, contents] of Object.entries({ previous, current })) {
  const bundle = await build({
    entryPoints: [path],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'CalculationBenchmark',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /core\/src\/async-calculation\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/core/src'),
          }));
        },
      },
    ],
  });
  sources[variant] =
    bundle.outputFiles[0].text +
    `
self.onmessage = async ({data}) => {
  try {
    let sourceReads = 0;
    const calc = new CalculationBenchmark.AsyncCalculation(async (sheet, row, column) => {
      if (sheet === 'source') { sourceReads++; return { rowId: row, columnId: column, input: { type:'number', value:1 } }; }
      return { rowId:row, columnId:column, input:{ type:'formula', expression:'SUM(Source!$A$1:$A$32)+' + row } };
    }, (name) => name === 'Source' ? 'source' : undefined);
    const read = calc.readSession();
    const start = performance.now();
    for (let row = 0; row < data.rows; row++) {
      const value = await read('formula', row, 0);
      if (value !== 32 + row) throw new Error('Wrong output ' + row);
    }
    const ms = performance.now() - start;
    const expectedReads = 32;
    if (sourceReads !== expectedReads) throw new Error('Wrong source read count ' + sourceReads);
    postMessage({ rows:data.rows, ms, sourceReads, checkedOutputs:data.rows });
  } catch (error) { postMessage({error:String(error)}); }
};`;
  hashes[variant] = createHash('sha256').update(sources[variant]).digest('hex');
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless dedicated Worker synthetic AsyncCalculation benchmark. Six alternating pairs per browser and 128/512 formulas over the same 32 scalar references; fresh calculator each sample, all outputs and exact source-read counts verified. Current reconstructs the experimental memo-first lookup order; previous uses the retained source. Excludes IndexedDB, real workbook scheduling, UI, cancellation, peak memory, reference devices and end-to-end latency.',
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
          async ({ sources, round }) => {
            const urls = [],
              workers = {};
            for (const [name, source] of Object.entries(sources)) {
              const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' }));
              urls.push(url);
              workers[name] = new Worker(url);
            }
            try {
              const samples = [];
              for (const variant of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
                const worker = workers[variant];
                for (const rows of round % 2 ? [512, 128] : [128, 512]) {
                  const sample = await new Promise((resolve, reject) => {
                    worker.onmessage = ({ data }) =>
                      data.error ? reject(new Error(data.error)) : resolve(data);
                    worker.onerror = (error) => reject(new Error(error.message));
                    worker.postMessage({ rows });
                  });
                  samples.push({ variant, ...sample });
                }
              }
              return samples;
            } finally {
              for (const worker of Object.values(workers)) worker.terminate();
              for (const url of urls) URL.revokeObjectURL(url);
            }
          },
          { sources, round },
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
