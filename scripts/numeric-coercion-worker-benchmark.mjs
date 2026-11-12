import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output)
  throw new Error('Usage: node scripts/numeric-coercion-worker-benchmark.mjs NEW_REPORT');
const path = resolve('packages/formula/src/coercion.ts');
const current = readFileSync(path, 'utf8');
const guard = `  if (typeof v === 'string') {
    const text = v.trimStart();
    if (text[0] === '0') {
      const radix = text.charCodeAt(1) | 32;
      if (radix === 120 || radix === 98 || radix === 111) fail();
    }
  }
`;
if (current.split(guard).length !== 2) throw new Error('Cannot reconstruct previous guard');
const previous = current.replace(guard, '');
const regex = current.replace(
  guard,
  "  if (typeof v === 'string' && /^0[xbo]/i.test(v.trim())) fail();\n",
);
const scripts = {},
  hashes = {};
for (const [variant, contents] of Object.entries({ previous, regex, current })) {
  const result = await build({
    entryPoints: [path],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'Coercion',
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /formula\/src\/coercion\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/formula/src'),
          }));
        },
      },
    ],
  });
  scripts[variant] =
    result.outputFiles[0].text +
    `
self.onmessage = ({data}) => {
  const values = data.text ? ['7', '  -2.5 ', '1e2', '0'] : [7, -2.5, 100, 0];
  let sum = 0;
  const start = performance.now();
  for (let i = 0; i < 1000000; i++) sum += Coercion.number(values[i % 4]);
  const ms = performance.now() - start;
  if (sum !== 26125000) throw new Error('Incorrect coercion sum: ' + sum);
  postMessage({ms, sum});
};`;
  hashes[variant] = createHash('sha256').update(scripts[variant]).digest('hex');
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless browser Worker microbenchmark of one million valid scalar numeric or decimal-text coercions. Six alternating fresh-Worker triples per browser; previous removes the radix guard and regex reconstructs the original guard. Sum verified each sample. Excludes formulas, invalid inputs, storage, rendering, peak memory and other devices.',
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
      const samples = await page.evaluate(
        async ({ scripts, round }) => {
          const result = [];
          for (const variant of round % 2
            ? ['current', 'regex', 'previous']
            : ['previous', 'regex', 'current']) {
            const url = URL.createObjectURL(
              new Blob([scripts[variant]], { type: 'text/javascript' }),
            );
            const worker = new Worker(url);
            try {
              for (const text of round % 2 ? [true, false] : [false, true]) {
                const sample = await new Promise((resolve, reject) => {
                  worker.onmessage = ({ data }) => resolve(data);
                  worker.onerror = reject;
                  worker.postMessage({ text });
                });
                result.push({ variant, text, ...sample });
              }
            } finally {
              worker.terminate();
              URL.revokeObjectURL(url);
            }
          }
          return result;
        },
        { scripts, round },
      );
      result.rounds.push({ round: round + 1, samples });
      save();
    }
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
