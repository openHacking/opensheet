import { build } from 'esbuild';
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { cpus, release } from 'node:os';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/stream-chunk-benchmark.mjs NEW_REPORT');
const bundles = {};
for (const variant of ['scheduled', 'noYield', 'timerYield']) {
  const result = await build({
    stdin: {
      contents: `export { readBinary, writeBinary } from './packages/core/src/binary-file.ts';
export { parseWorkbookJSON } from './packages/core/src/json-stream.ts';
export { createSnapshot } from './packages/core/src/model.ts';
export { encodeWorkbookFile } from './packages/core/src/file-codec.ts';`,
      resolveDir: process.cwd(),
      loader: 'ts',
    },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: `Stream_${variant}`,
    write: false,
    plugins:
      variant !== 'scheduled'
        ? [
            {
              name: 'disable-task-yield',
              setup(build) {
                build.onLoad({ filter: /\/(binary-file|json-stream)\.ts$/ }, ({ path }) => {
                  const text = readFileSync(path, 'utf8');
                  const target = 'await yieldToTask();';
                  if (text.split(target).length !== 2)
                    throw new Error(`Unexpected yield code in ${path}`);
                  return {
                    contents: text.replace(
                      target,
                      variant === 'noYield'
                        ? 'await Promise.resolve();'
                        : 'await new Promise<void>((resolve) => setTimeout(resolve, 0));',
                    ),
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
    'Synthetic in-page parser benchmark. The scheduled variant uses current production task yields. The timerYield control uses setTimeout; noYield uses microtasks and cannot provide the same cancellation guarantee. Batches rotate order and verify JSON payload and binary metadata. Excludes Worker transfer, persistence, large workbooks, peak memory and reference devices.',
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
      if (!globalThis.crypto.randomUUID) {
        let id = 0;
        Object.defineProperty(globalThis.crypto, 'randomUUID', {
          value: () => `00000000-0000-4000-8000-${String(++id).padStart(12, '0')}`,
        });
      }
      const variants = {
        scheduled: globalThis.Stream_scheduled,
        noYield: globalThis.Stream_noYield,
        timerYield: globalThis.Stream_timerYield,
      };
      const payload = 'x'.repeat(16000);
      const json = new TextEncoder().encode(JSON.stringify({ payload }));
      const api = variants.scheduled;
      const metadata = api.encodeWorkbookFile(
        api.createSnapshot({ sheets: [{ name: 'Bench', rows: 10, columns: 4 }] }),
      );
      const pieces = [];
      for await (const bytes of api.writeBinary(metadata, (async function* () {})()))
        pieces.push(bytes);
      const binary = new Uint8Array(await new Blob(pieces).arrayBuffer());
      const results = [];
      for (const format of ['json', 'binary'])
        for (const size of [1, 64, 65536]) {
          const bytes = format === 'json' ? json : binary;
          const iterations = size === 1 ? 5 : 20;
          for (let round = 0; round < 6; round++)
            for (const variant of [
              ['scheduled', 'noYield', 'timerYield'],
              ['timerYield', 'scheduled', 'noYield'],
              ['noYield', 'timerYield', 'scheduled'],
            ][round % 3]) {
              const start = performance.now();
              for (let iteration = 0; iteration < iterations; iteration++) {
                let offset = 0,
                  root;
                const source = new ReadableStream({
                  pull(controller) {
                    if (offset >= bytes.length) controller.close();
                    else {
                      controller.enqueue(bytes.subarray(offset, offset + size));
                      offset += size;
                    }
                  },
                });
                const parser =
                  format === 'json'
                    ? variants[variant].parseWorkbookJSON(source)
                    : variants[variant].readBinary(source);
                for await (const item of parser) if ('metadata' in item) root = item.metadata;
                if (
                  format === 'json'
                    ? root?.payload !== payload
                    : root?.sheets?.[0]?.name !== 'Bench'
                )
                  throw new Error(`Bad ${format} result: ${variant}`);
              }
              results.push({
                format,
                chunkBytes: size,
                iterations,
                round,
                variant,
                ms: performance.now() - start,
                bytesPerInput: bytes.length,
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
