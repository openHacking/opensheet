import { build } from 'esbuild';
import { chromium } from '@playwright/test';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/capacity-search-paired.mjs NEW_REPORT.json');
const enginePath = resolve('packages/core/src/engine.ts');
const current = readFileSync(enginePath, 'utf8');
const stagedStart = current.indexOf(
  '    const page = await this.stagedPage(address, head, local);',
);
const stagedEnd = current.indexOf(
  '  /** Internal bridge; caller must keep head and metadata immutable',
  stagedStart,
);
const searchStart = current.indexOf('    // A zero-cache search would otherwise reread');
const searchEnd = current.indexOf('    let wrap: number | undefined,', searchStart);
if (stagedStart < 0 || stagedEnd < 0 || searchStart < 0 || searchEnd < 0)
  throw Error('Cannot reconstruct prior search');
const priorCell = `    let page = local?.address === address ? local.page : undefined;
    if (!page) {
      page = await this.block(address, head);
      if (local) {
        local.address = page.bytes <= CONDITIONAL_READ_PAGE_BUDGET ? address : undefined;
        local.page = page.bytes <= CONDITIONAL_READ_PAGE_BUDGET ? page : undefined;
      }
    }
    return page.cell((sh.rowOrder[row] % 64) * 32 + (sh.columnOrder[column] % 32));
  }
`;
let control = current.slice(0, stagedStart) + priorCell + current.slice(stagedEnd);
const adjustedStart = control.indexOf('    // A zero-cache search would otherwise reread');
const adjustedEnd = control.indexOf('    let wrap: number | undefined,', adjustedStart);
control = control.slice(0, adjustedStart) + control.slice(adjustedEnd);
const call = '? await calculate(sheetId, row, column)';
if (control.split(call).length !== 2) throw Error('Cannot reconstruct prior calculator call');
control = control.replace(call, '? await this.calculator.value(sheetId, row, column)');
const sha256 = (data) => createHash('sha256').update(data).digest('hex');
let powerMode = 'unavailable';
try {
  powerMode = execFileSync('pmset', ['-g', 'batt'], { encoding: 'utf8' }).trim();
} catch {}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  build: localBuildFingerprint(),
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  node: process.version,
  powerMode,
  samples: [],
  bundles: {},
  scope:
    'Two alternating fresh Chromium contexts per variant, local production-preview UI with its engine Worker replaced by esbuild bundles of current source or reconstructed pre-search-page-reuse source. Each sample runs the 1m three-round lab including data, viewport and history checks. Worker bundling differs from the production Vite asset in both variants. No native input, process peak memory, other browsers/devices or independent release-commit acceptance.',
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
let browser;
try {
  const bundles = {};
  for (const [name, code] of [
    ['control', control],
    ['current', current],
  ]) {
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
    bundles[name] = result.outputFiles[0].contents;
    report.bundles[name] = sha256(bundles[name]);
  }
  browser = await chromium.launch();
  report.browser = browser.version();
  for (let round = 1; round <= 2; round++)
    for (const variant of round % 2 ? ['control', 'current'] : ['current', 'control']) {
      const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
      let page,
        workerRequests = 0;
      const errors = [],
        failedRequests = [];
      try {
        page = await context.newPage();
        page.on('pageerror', (error) => errors.push(String(error)));
        page.on('console', (message) => {
          if (message.type() === 'error') errors.push(message.text());
        });
        await page.addInitScript(() => {
          window.workerDiagnostics = [];
          const OriginalWorker = window.Worker;
          window.Worker = class extends OriginalWorker {
            constructor(...args) {
              super(...args);
              this.addEventListener('error', (event) =>
                window.workerDiagnostics.push({ kind: 'error', message: event.message }),
              );
              this.addEventListener('message', (event) => {
                if (event.data?.error)
                  window.workerDiagnostics.push({ kind: 'response', error: event.data.error });
              });
            }
          };
        });
        page.on('requestfailed', (request) =>
          failedRequests.push({ url: request.url(), failure: request.failure()?.errorText }),
        );
        await page.route('**/assets/engine.worker-*.js', (route) => {
          workerRequests++;
          return route.fulfill({
            status: 200,
            contentType: 'text/javascript',
            body: Buffer.from(bundles[variant]),
          });
        });
        await page.goto('http://127.0.0.1:59849/#performance');
        await page
          .locator('#scene-title')
          .filter({ hasText: 'Performance Lab' })
          .waitFor({ timeout: 10000 });
        await page.waitForFunction(() => !document.querySelector('#scene-select').disabled);
        await page.locator('#perf-budget').fill('256');
        await page.locator('#perf-cells').fill('1000000');
        await page.locator('#perf-run').click();
        await page
          .locator('#perf-progress')
          .filter({ hasText: 'Finished' })
          .waitFor({ timeout: 3600000 });
        const downloadPromise = page.waitForEvent('download');
        await page.locator('#perf-download').click();
        const file = await (await downloadPromise).path(),
          result = JSON.parse(readFileSync(file, 'utf8'));
        if (!workerRequests || errors.length)
          throw Error(
            `Worker interception or page errors: ${JSON.stringify({ workerRequests, errors })}`,
          );
        const run = result.results?.[0];
        if (
          !run ||
          run.cells !== 1000000 ||
          (!run.passed && run.failure !== 'Interaction threshold exceeded')
        )
          throw Error(
            `Unexpected result: ${JSON.stringify({ passed: run?.passed, failure: run?.failure, cells: run?.cells })}`,
          );
        report.samples.push({ round, variant, workerRequests, errors, result });
        save();
        console.log(
          JSON.stringify({
            round,
            variant,
            passed: run.passed,
            failure: run.failure,
            scroll: run.rounds.map((x) => x.scrollP95Ms),
            edit: run.rounds.map((x) => x.editP95Ms),
          }),
        );
      } catch (error) {
        report.samples.push({
          round,
          variant,
          error: String(error),
          workerRequests,
          errors,
          failedRequests,
          workerDiagnostics: page
            ? await page.evaluate(() => window.workerDiagnostics).catch(() => [])
            : [],
          body: page
            ? (
                await page
                  .locator('body')
                  .innerText()
                  .catch(() => '')
              ).slice(0, 3000)
            : '',
        });
        save();
        throw error;
      } finally {
        await context.close();
      }
    }
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
} finally {
  await browser?.close();
}
