import { chromium, firefox, webkit } from '@playwright/test';
import { build as bundle } from 'esbuild';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, readdirSync, unlinkSync, writeFileSync } from 'node:fs';
import { arch, cpus, platform, totalmem } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const option = (name, fallback) =>
  process.argv.find((arg) => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback;
const output = option('output', ''),
  engine = option('browser', 'chromium'),
  count = Number(option('cells', '100000')),
  rounds = Number(option('rounds', '3'));
const launchers = { chromium, firefox, webkit };
if (
  !output ||
  existsSync(output) ||
  !launchers[engine] ||
  !Number.isSafeInteger(count) ||
  count < 10000 ||
  count > 1000000 ||
  count % 10000 ||
  !Number.isSafeInteger(rounds) ||
  rounds < 1 ||
  rounds > 10
)
  throw new Error('Specify a new output, valid browser, and 10k–1m cells in 10k steps');
const path = 'apps/playground/dist/xlsx-export-benchmark.js';
if (existsSync(path)) throw new Error('Refusing to overwrite benchmark bundle');
const workers = readdirSync('apps/playground/dist/assets').filter((name) =>
  /^xlsx-export\.worker-.*\.js$/.test(name),
);
if (workers.length !== 1)
  throw new Error('Build the playground with exactly one XLSX export Worker');
const source = sourceFingerprint(),
  build = localBuildFingerprint();
let browser;
try {
  await bundle({
    entryPoints: ['scripts/xlsx-export-benchmark-entry.ts'],
    bundle: true,
    platform: 'browser',
    format: 'esm',
    minify: true,
    outfile: path,
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
    async ({ count, rounds: roundCount, workerPath }) => {
      const { write, compare } = await import('/xlsx-export-benchmark.js');
      const book = window.opensheet.getWorkbook();
      await book.generate(count);
      await book.getSheets()[0].setFreeze(1, 1);
      const measure = async (action) => {
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
          const value = await action();
          const elapsedMs = performance.now() - start;
          await new Promise((resolve) =>
            requestAnimationFrame(() => requestAnimationFrame(resolve)),
          );
          return { value, elapsedMs, maxFrameMs, frames };
        } finally {
          running = false;
        }
      };
      const snapshot = await measure(() => book.toJSON());
      const file = snapshot.value,
        rounds = [];
      const main = () => write(file);
      const worker = () =>
        new Promise((resolve, reject) => {
          const w = new Worker(workerPath, { type: 'module' });
          w.onmessage = (event) => {
            w.terminate();
            if (event.data.error) reject(new Error(event.data.error));
            else resolve({ bytes: new Uint8Array(event.data.buffer), report: event.data.report });
          };
          w.onerror = () => {
            w.terminate();
            reject(new Error('Worker failed'));
          };
          w.postMessage(file);
        });
      for (let i = 0; i < roundCount; i++) {
        const order = i % 2 ? ['worker', 'main'] : ['main', 'worker'],
          samples = {};
        for (const name of order) samples[name] = await measure(name === 'main' ? main : worker);
        if (
          JSON.stringify(samples.main.value.report) !== JSON.stringify(samples.worker.value.report)
        )
          throw new Error('Reports differ');
        const verified = compare(samples.main.value.bytes, samples.worker.value.bytes);
        if (verified.cells !== count)
          throw new Error(`Expected ${count} cells, got ${verified.cells}`);
        const freezeDropped = samples.main.value.report.issues.some(
          (issue) => issue.code === 'VIEW_NOT_EXPORTED',
        );
        if (verified.frozen === freezeDropped)
          throw new Error('Freeze and compatibility report disagree');
        rounds.push({
          order,
          main: {
            elapsedMs: samples.main.elapsedMs,
            maxFrameMs: samples.main.maxFrameMs,
            frames: samples.main.frames,
            bytes: samples.main.value.bytes.byteLength,
          },
          worker: {
            elapsedMs: samples.worker.elapsedMs,
            maxFrameMs: samples.worker.maxFrameMs,
            frames: samples.worker.frames,
            bytes: samples.worker.value.bytes.byteLength,
          },
          verified,
          report: samples.main.value.report.summary,
        });
      }
      return {
        cells: count,
        snapshot: {
          elapsedMs: snapshot.elapsedMs,
          maxFrameMs: snapshot.maxFrameMs,
          frames: snapshot.frames,
        },
        rounds,
      };
    },
    { count, rounds, workerPath: `/assets/${workers[0]}` },
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
      'Production Worker versus a separate minified main-thread bundle, same snapshot and alternating order in one context. Timed Worker includes startup and structured clone; snapshot measured separately. Compares every parsed cell input and the complete first worksheet XML after timings. Excludes UI dialog/download, process memory, physical power configuration and cross-device acceptance.',
    result,
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(JSON.stringify(report));
} catch (error) {
  if (!existsSync(output))
    writeFileSync(
      output,
      JSON.stringify(
        {
          date: new Date().toISOString(),
          engine,
          cells: count,
          source,
          build,
          failure: String(error),
        },
        null,
        2,
      ) + '\n',
      { flag: 'wx' },
    );
  throw error;
} finally {
  await browser?.close();
  if (existsSync(path)) unlinkSync(path);
}
