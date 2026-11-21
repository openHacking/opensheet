import { build } from 'esbuild';
import { chromium } from '@playwright/test';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/history-large-paired.mjs NEW_REPORT.json');
const baselineCommit = 'da9ade8';
const changed = execFileSync(
  'rtk',
  ['proxy', 'git', 'diff', '--name-only', baselineCommit, 'HEAD', '--', 'packages'],
  { encoding: 'utf8' },
).trim();
if (changed !== 'packages/core/src/engine.ts')
  throw new Error(`Unexpected baseline differences: ${changed}`);
const variants = {
  baseline: execFileSync(
    'rtk',
    ['proxy', 'git', 'show', `${baselineCommit}:packages/core/src/engine.ts`],
    { encoding: 'utf8', maxBuffer: 5 * 1024 * 1024 },
  ),
  current: readFileSync('packages/core/src/engine.ts', 'utf8'),
};
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
let powerMode = 'unavailable';
try {
  powerMode = execFileSync('pmset', ['-g', 'batt'], { encoding: 'utf8' }).trim();
} catch {}
const report = {
  startedAt: new Date().toISOString(),
  baselineCommit,
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length },
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  powerMode,
  bundles: {},
  samples: [],
  scope:
    'Three alternating fresh Chromium contexts per variant using production-preview UI and substituted esbuild Workers. Each generates 1m mixed cells, makes 100 single-cell edits, measures three reopen operations, a reopen halfway through undo and every undo/redo, checks each changed value and verifies all 1m cells after history. Undo/redo timings exclude the verification reads. This measures one local workload, not metadata-heavy histories, peak memory, cold OS caches, other devices or Vite Worker delivery.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
let browser;
try {
  const bundles = {};
  for (const [variant, code] of Object.entries(variants)) {
    const result = await build({
      entryPoints: [resolve('packages/core/src/engine.worker.js')],
      bundle: true,
      platform: 'browser',
      format: 'esm',
      target: 'es2022',
      write: false,
      minify: true,
      plugins: [
        {
          name: 'engine-variant',
          setup(builder) {
            builder.onLoad({ filter: /core\/src\/engine\.ts$/ }, () => ({
              contents: code,
              loader: 'ts',
              resolveDir: resolve('packages/core/src'),
            }));
          },
        },
      ],
    });
    bundles[variant] = result.outputFiles[0].contents;
    report.bundles[variant] = {
      sourceSha256: sha256(code),
      bundleSha256: sha256(bundles[variant]),
    };
  }
  browser = await chromium.launch();
  report.browser = browser.version();
  for (let round = 1; round <= 3; round++) {
    for (const variant of round % 2 ? ['baseline', 'current'] : ['current', 'baseline']) {
      const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
      let workerRequests = 0;
      try {
        const page = await context.newPage();
        const errors = [];
        page.on('pageerror', (error) => errors.push(String(error)));
        await page.route('**/assets/engine.worker-*.js', (route) => {
          workerRequests++;
          return route.fulfill({
            status: 200,
            contentType: 'text/javascript',
            body: Buffer.from(bundles[variant]),
          });
        });
        await page.goto('http://127.0.0.1:5186/#budget');
        await page.waitForFunction(() => !!window.opensheet?.selection.get());
        const data = await page.evaluate(async () => {
          const app = window.opensheet;
          let book = app.getWorkbook();
          await book.generate(1_000_000);
          let sheet = book.getSheets()[0];
          for (let step = 1; step <= 100; step++) await sheet.range('A1').setValues([[step]]);
          const id = book.id,
            revision = book.revision,
            reopenMs = [];
          for (let i = 0; i < 3; i++) {
            const start = performance.now();
            await app.open(id);
            reopenMs.push(performance.now() - start);
            book = app.getWorkbook();
            sheet = book.getSheets()[0];
            if (book.revision !== revision || (await sheet.range('A1').getValues())[0][0] !== 100)
              throw new Error('Reopened value or revision differs');
          }
          const undoMs = [],
            redoMs = [];
          let midpointReopenMs;
          for (let step = 99; step >= 0; step--) {
            const start = performance.now();
            const applied = await book.undo();
            undoMs.push(performance.now() - start);
            if (!applied || (await sheet.range('A1').getValues())[0][0] !== step)
              throw new Error(`Undo mismatch ${step}`);
            if (step === 50) {
              const midRevision = book.revision;
              const reopenStart = performance.now();
              await app.open(id);
              midpointReopenMs = performance.now() - reopenStart;
              book = app.getWorkbook();
              sheet = book.getSheets()[0];
              if (
                book.revision !== midRevision ||
                (await sheet.range('A1').getValues())[0][0] !== step
              )
                throw new Error('Midpoint reopen changed revision or value');
            }
          }
          for (let step = 1; step <= 100; step++) {
            const start = performance.now();
            const applied = await book.redo();
            redoMs.push(performance.now() - start);
            if (!applied || (await sheet.range('A1').getValues())[0][0] !== step)
              throw new Error(`Redo mismatch ${step}`);
          }
          const expected = (n) => {
            let x = (Math.imul(n + 1, 1664525) + 1013904223) >>> 0;
            if (n % 10 < 6) return n === 0 ? 100 : x / 1000;
            if (n % 10 === 9) return ((n / 10) | 0) % 2 === 1;
            let text = '';
            for (let i = 0; i < 32; i++) {
              x = (Math.imul(x, 1664525) + 1013904223) >>> 0;
              text += String.fromCharCode(33 + (x % 94));
            }
            return text;
          };
          let verified = 0;
          for (let start = 0; start < 10_000; start += 100) {
            const matrix = await sheet.range(`A${start + 1}:CV${start + 100}`).getValues();
            for (let row = 0; row < matrix.length; row++)
              for (let col = 0; col < 100; col++) {
                const n = (start + row) * 100 + col;
                if (matrix[row][col] !== expected(n)) throw new Error(`Value mismatch at ${n}`);
                verified++;
              }
          }
          const stats = await book.storageStats();
          await book.deleteStorage();
          return { reopenMs, midpointReopenMs, undoMs, redoMs, verified, stats };
        });
        if (!workerRequests || errors.length)
          throw new Error(`Worker requests: ${workerRequests}, page errors: ${errors.join('; ')}`);
        report.samples.push({ round, variant, workerRequests, ...data });
        save();
        console.log(
          JSON.stringify({
            round,
            variant,
            reopenMs: data.reopenMs,
            midpointReopenMs: data.midpointReopenMs,
            undoTotalMs: data.undoMs.reduce((a, b) => a + b, 0),
            redoTotalMs: data.redoMs.reduce((a, b) => a + b, 0),
            verified: data.verified,
          }),
        );
      } catch (error) {
        report.samples.push({ round, variant, workerRequests, error: String(error) });
        save();
        throw error;
      } finally {
        await context.close();
      }
    }
  }
} finally {
  await browser?.close();
}
