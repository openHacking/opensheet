import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/aggregate-worker-benchmark.mjs NEW_REPORT');
const path = resolve('packages/formula/src/aggregate.ts'),
  current = readFileSync(path, 'utf8');
const start = current.indexOf("    if (typeof value === 'number') {");
const end = current.indexOf("    if (this.name === 'AND' || this.name === 'OR')", start);
if (start < 0 || end < 0) throw new Error('Cannot reconstruct previous accumulator');
const previous =
  current.slice(0, start) +
  `    if (typeof value === 'number') {
      this.sum += value; this.count++; this.min = Math.min(this.min, value); this.max = Math.max(this.max, value);
    } else if (this.name === 'COUNT' && !referenced && (typeof value === 'boolean' || (typeof value === 'string' && /^[+-]?(?:\\d+(?:\\.\\d*)?|\\.\\d+)(?:[eE][+-]?\\d+)?$/.test(value.trim()) && Number.isFinite(Number(value))))) this.count++;
` +
  current.slice(end);
const scripts = {},
  hashes = {};
for (const [name, contents] of Object.entries({ previous, current })) {
  const result = await build({
    entryPoints: [path],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'AggregateBenchmark',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /formula\/src\/aggregate\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/formula/src'),
          }));
        },
      },
    ],
  });
  scripts[name] =
    result.outputFiles[0].text +
    `\nself.onmessage = ({data}) => {
    const start = performance.now(); let value = 0;
    for (let pass = 0; pass < 20; pass++) {
      const accumulator = new AggregateBenchmark.Aggregate(data.name);
      for (let i = 0; i < 100000; i++) accumulator.add(data.mixed ? (i % 4 === 0 ? 1 : i % 4 === 1 ? '7' : i % 4 === 2 ? true : null) : 1, true);
      const result = accumulator.result();
      if (result !== data.expected) throw new Error('Incorrect result in pass ' + pass);
      value += result;
    }
    const ms = performance.now() - start;
    postMessage({name:data.name, mixed:data.mixed, value, ms});
  };`;
  hashes[name] = createHash('sha256').update(scripts[name]).digest('hex');
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless dedicated browser Worker microbenchmark; source-bundled current and reconstructed previous accumulator, not a full workbook. Each sample performs 20 passes over 100,000 referenced values (all numeric or repeating number/text/boolean/blank) into SUM/MIN/MAX/COUNT and verifies each result. Six alternating rounds in each fresh Worker pair. Excludes formula parsing, dependencies, IndexedDB, output/UI latency, worker/process memory, other devices and direct argument behavior.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch(),
    result = { engine, version: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    const context = await browser.newContext();
    try {
      const page = await context.newPage();
      for (let round = 1; round <= 6; round++) {
        const entry = { round, samples: [] };
        result.rounds.push(entry);
        try {
          entry.samples = await page.evaluate(
            async ({ scripts, round }) => {
              const workers = {},
                urls = [];
              try {
                for (const [name, script] of Object.entries(scripts)) {
                  const url = URL.createObjectURL(new Blob([script], { type: 'text/javascript' }));
                  urls.push(url);
                  workers[name] = new Worker(url);
                }
                const samples = [];
                for (const variant of round % 2
                  ? ['previous', 'current']
                  : ['current', 'previous']) {
                  for (const mixed of round % 2 ? [false, true] : [true, false])
                    for (const name of ['SUM', 'MIN', 'MAX', 'COUNT']) {
                      const expected =
                        name === 'SUM' || name === 'COUNT' ? (mixed ? 25000 : 100000) : 1;
                      const response = await new Promise((resolve, reject) => {
                        workers[variant].onmessage = ({ data }) => resolve(data);
                        workers[variant].onerror = reject;
                        workers[variant].postMessage({ name, mixed, expected });
                      });
                      if (response.value !== expected * 20)
                        throw new Error(`${variant}/${name}/${mixed}: ${response.value}`);
                      samples.push({ variant, ...response, verified: 2000000 });
                    }
                }
                return samples;
              } finally {
                Object.values(workers).forEach((worker) => worker.terminate());
                urls.forEach((url) => URL.revokeObjectURL(url));
              }
            },
            { scripts, round },
          );
          console.log(JSON.stringify({ engine, round, samples: entry.samples }));
        } catch (error) {
          entry.error = String(error);
          throw error;
        } finally {
          save();
        }
      }
    } finally {
      await context.close();
    }
  } finally {
    save();
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
