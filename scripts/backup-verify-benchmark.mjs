import { chromium, firefox, webkit } from '@playwright/test';
import { build as bundle } from 'esbuild';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync, unlinkSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { cpus, totalmem, platform, arch } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';
const option = (key, fallback) =>
  process.argv.find((s) => s.startsWith(`--${key}=`))?.slice(key.length + 3) ?? fallback;
const output = option('output', ''),
  engine = option('browser', 'chromium'),
  count = Number(option('cells', '1000000')),
  profile = option('profile', 'false') === 'true';
const launchers = { chromium, firefox, webkit };
if (
  !output ||
  existsSync(output) ||
  !launchers[engine] ||
  !Number.isSafeInteger(count) ||
  count < 10000 ||
  count > 10000000 ||
  count % 10000
)
  throw new Error('Specify a new output, valid browser and cell count');
const path = 'apps/playground/dist/backup-verify-benchmark.js';
if (existsSync(path)) throw new Error('Refusing to overwrite benchmark bundle');
const workerAssets = readdirSync('apps/playground/dist/assets').filter((name) =>
  /^backup-verifier\.worker-.*\.js$/.test(name),
);
if (workerAssets.length !== 1)
  throw new Error('Build the playground with exactly one backup verification Worker');
const workerPath = `/assets/${workerAssets[0]}`;
const source = sourceFingerprint(),
  build = localBuildFingerprint();
let browser;
try {
  await bundle({
    entryPoints: ['scripts/backup-verify-entry.ts'],
    bundle: true,
    platform: 'browser',
    format: 'esm',
    minify: true,
    outfile: path,
    alias: { '@opensheetjs/formula': resolve('packages/formula/src/index.ts') },
    logLevel: 'silent',
  });
  const bundleHash = createHash('sha256').update(readFileSync(path)).digest('hex');
  browser = await launchers[engine].launch();
  const page = await browser.newPage();
  await page.goto('http://127.0.0.1:4174/#performance');
  await page.waitForFunction(
    () =>
      document.querySelector('#scene-title')?.textContent === 'Performance Lab' &&
      !document.querySelector('#scene-select')?.disabled,
  );
  const result = await page.evaluate(
    async ({ count, profile, workerPath }) => {
      const { verifyBinary, verifyJSON, compareBinaryBackup, profileBinaryVerification } =
        await import('/backup-verify-benchmark.js');
      const book = window.opensheet.getWorkbook();
      await book.generate(count);
      const jsonParts = [],
        binaryParts = [];
      for await (const part of book.streamJSON()) jsonParts.push(part);
      for await (const part of book.streamBinary()) binaryParts.push(part);
      const json = new Blob(jsonParts),
        binary = new Blob(binaryParts);
      const measure = async (
        action,
        validate = (summary) =>
          summary.cells === count && summary.sheets === 1 && summary.workbookId === book.id,
      ) => {
        let running = true,
          last,
          maxFrameMs = 0,
          frames = 0;
        const tick = (time) => {
          if (last !== undefined) {
            maxFrameMs = Math.max(maxFrameMs, time - last);
            frames++;
          }
          last = time;
          if (running) requestAnimationFrame(tick);
        };
        await new Promise((resolve) =>
          requestAnimationFrame((time) => {
            last = time;
            requestAnimationFrame(tick);
            resolve();
          }),
        );
        const start = performance.now();
        try {
          const summary = await action();
          const elapsedMs = performance.now() - start;
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          if (!validate(summary)) throw new Error('Backup verification summary differs');
          return { elapsedMs, maxFrameMs, frames, summary };
        } finally {
          running = false;
        }
      };
      return {
        cells: count,
        jsonBytes: json.size,
        binaryBytes: binary.size,
        ...(profile
          ? {
              idle: await measure(
                () => new Promise((resolve) => setTimeout(() => resolve(true), 1000)),
                (summary) => summary === true,
              ),
            }
          : {}),
        json: await measure(() => verifyJSON(json)),
        binary: await measure(() => verifyBinary(binary)),
        workerBinary: await measure(
          () =>
            new Promise((resolve, reject) => {
              const worker = new Worker(workerPath, { type: 'module' });
              worker.onmessage = (event) => {
                worker.terminate();
                if (event.data.error !== undefined) reject(new Error(event.data.error));
                else resolve(event.data.summary);
              };
              worker.onerror = (event) => {
                worker.terminate();
                reject(new Error(event.message));
              };
              worker.postMessage({ file: new File([binary], 'backup.opensheet'), json: false });
            }),
        ),
        comparison: await measure(
          () => compareBinaryBackup(book, binary),
          (summary) => summary.matches && summary.revision === book.revision,
        ),
        ...(profile
          ? {
              profile: await measure(
                () => profileBinaryVerification(binary),
                (summary) => summary.cells === count,
              ),
            }
          : {}),
      };
    },
    { count, profile, workerPath },
  );
  const report = {
    date: new Date().toISOString(),
    engine,
    browser: browser.version(),
    source,
    build,
    bundleHash,
    host: { platform: platform(), arch: arch(), cpu: cpus()[0].model, memoryBytes: totalmem() },
    scope:
      'Production app, separate minified source verifier bundle on the main page and the production backup Worker; one fresh context, complete native backups of deterministic fixture, verification/comparison summaries and animation-frame intervals. Optional profile uses an instrumented replica and read timings include async waits. Excludes fixture and initial export from timings; comparison includes its own export. No Worker/process memory, controlled power mode or cross-device acceptance.',
    result,
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify(report));
} finally {
  await browser?.close();
  if (existsSync(path)) unlinkSync(path);
}
