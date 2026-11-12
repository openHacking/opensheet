import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, release } from 'node:os';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/json-decode-slices-benchmark.mjs NEW_REPORT');
const bundles = {};
for (const variant of ['bounded', 'whole']) {
  const result = await build({
    entryPoints: ['packages/core/src/json-stream.ts'],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: `Parser_${variant}`,
    write: false,
    plugins:
      variant === 'whole'
        ? [
            {
              name: 'whole-chunk-control',
              setup(build) {
                build.onLoad({ filter: /\/json-stream\.ts$/ }, ({ path }) => {
                  const text = readFileSync(path, 'utf8');
                  const target = 'const end = Math.min(inputOffset + 65536, input!.length);';
                  if (text.split(target).length !== 2)
                    throw new Error('Cannot reconstruct control');
                  return {
                    contents: text.replace(target, 'const end = input!.length;'),
                    loader: 'ts',
                  };
                });
              },
            },
          ]
        : [],
  });
  bundles[variant] = result.outputFiles[0].text;
}
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  bundleSha256: Object.fromEntries(
    Object.entries(bundles).map(([k, v]) => [k, createHash('sha256').update(v).digest('hex')]),
  ),
  scope:
    'Paired synthetic in-page JSON parser benchmark, with a reconstructed whole-chunk decode control. Five alternating rounds per engine and chunk size verify every decoded value. The control lacks bounded per-decode allocation; neither variant is an old commit. Excludes Worker transfer, persisted import, process memory and reference devices.',
  browsers: {},
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [name, engine] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await engine.launch({ headless: true });
  try {
    const page = await browser.newPage();
    for (const content of Object.values(bundles)) await page.addScriptTag({ content });
    const samples = await page.evaluate(async () => {
      const variants = { bounded: globalThis.Parser_bounded, whole: globalThis.Parser_whole };
      const values = Array.from({ length: 10000 }, (_, i) => `${i}:${'中😀abc'.repeat(8)}`);
      const bytes = new TextEncoder().encode(JSON.stringify({ values }));
      const results = [];
      for (const size of [1024, 65536, bytes.length])
        for (let round = 0; round < 5; round++) {
          for (const variant of round % 2 ? ['whole', 'bounded'] : ['bounded', 'whole']) {
            let offset = 0,
              root;
            const stream = new ReadableStream({
              pull(controller) {
                if (offset === bytes.length) controller.close();
                else {
                  controller.enqueue(bytes.subarray(offset, Math.min(bytes.length, offset + size)));
                  offset = Math.min(bytes.length, offset + size);
                }
              },
            });
            const start = performance.now();
            for await (const part of variants[variant].parseWorkbookJSON(stream))
              if ('metadata' in part) root = part.metadata;
            const ms = performance.now() - start;
            if (
              root?.values?.length !== values.length ||
              values.some((value, i) => root.values[i] !== value)
            )
              throw new Error(`${variant}: decoded values differ`);
            results.push({
              chunkBytes: size,
              inputBytes: bytes.length,
              round,
              variant,
              ms,
              verifiedValues: values.length,
            });
          }
        }
      return results;
    });
    report.browsers[name] = { version: browser.version(), samples };
    save();
    console.log(name, 'completed', samples.length, 'samples');
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
