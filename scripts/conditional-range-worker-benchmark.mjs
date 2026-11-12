import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output)
  throw new Error('Usage: node scripts/conditional-range-worker-benchmark.mjs NEW_REPORT');
const path = resolve('packages/formula/src/ranges.ts');
const source = readFileSync(path, 'utf8');
const start = source.indexOf('export const cell =');
if (start < 0) throw new Error('Cannot reconstruct prior cell helper');
const previous =
  source.slice(0, start) +
  `export const cell = (range: Area, row: number, column: number): Reference => ({
    ...range.start, row: range.start.row + row, column: range.start.column + column
  });\n`;
const scripts = {},
  hashes = {};
for (const [name, contents] of Object.entries({ previous, current: source })) {
  const result = await build({
    stdin: {
      contents:
        "export { conditional } from './packages/formula/src/conditional.ts'; export { parseFormula } from './packages/formula/src/parser.ts';",
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'Bench',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /formula\/src\/ranges\.ts$/ }, () => ({
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
    `
self.onmessage = () => {
  const ast = Bench.parseFormula('SUMIF(A1:A100000,">0",B1)');
  let checksum = 0; const start = performance.now();
  for (let pass = 0; pass < 5; pass++) {
    const iterator = Bench.conditional(ast.name, ast.args); let step = iterator.next();
    while (!step.done) step = iterator.next(step.value.kind === 'literal' ? step.value.value : 1);
    if (step.value !== 100000) throw new Error('Wrong result ' + step.value);
    checksum += step.value;
  }
  postMessage({ms: performance.now() - start, checksum});
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
    'Dedicated headless browser Worker microbenchmark. Five 100,000-row SUMIF passes per sample; matching numeric synthetic reads and exact result checks. Six alternating rounds using fresh workers. Previous is a reconstructed helper; excludes real workbook reads, parser time, storage, UI, memory and other data or devices.',
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
    for (let round = 1; round <= 6; round++) {
      const entry = { round, samples: [] };
      result.rounds.push(entry);
      try {
        entry.samples = await page.evaluate(
          async ({ scripts, round }) => {
            const samples = [];
            for (const variant of round % 2 ? ['previous', 'current'] : ['current', 'previous']) {
              const url = URL.createObjectURL(
                new Blob([scripts[variant]], { type: 'text/javascript' }),
              );
              const worker = new Worker(url);
              try {
                const sample = await new Promise((resolve, reject) => {
                  worker.onmessage = ({ data }) => resolve(data);
                  worker.onerror = reject;
                  worker.postMessage(null);
                });
                samples.push({ variant, ...sample });
              } finally {
                worker.terminate();
                URL.revokeObjectURL(url);
              }
            }
            return samples;
          },
          { scripts, round },
        );
        console.log(JSON.stringify({ engine, ...entry }));
      } catch (error) {
        entry.error = String(error);
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
