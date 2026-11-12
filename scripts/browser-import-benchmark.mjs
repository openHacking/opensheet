import { chromium, firefox, webkit } from '@playwright/test';
import * as XLSX from 'xlsx';
import { createHash } from 'node:crypto';
import { existsSync, writeFileSync, readFileSync, readdirSync } from 'node:fs';
import { cpus, totalmem, platform, arch } from 'node:os';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';

const option = (name, fallback) =>
  process.argv.find((arg) => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback;
const output = option('output', ''),
  rows = Number(option('rows', '50000')),
  rounds = Number(option('rounds', '3'));
if (!output || existsSync(output)) throw new Error('Pass a new --output=report.json');
if (
  !Number.isSafeInteger(rows) ||
  rows < 100 ||
  rows > 100000 ||
  !Number.isSafeInteger(rounds) ||
  rounds < 1 ||
  rounds > 10
)
  throw new Error('Use 100–100000 rows and 1–10 rounds');
const engines = option('engines', 'chromium,webkit,firefox').split(',');
const profile = option('profile', 'false') === 'true';
const launchers = { chromium, firefox, webkit };
if (engines.some((name) => !launchers[name])) throw new Error('Unknown browser engine');
const values = Array.from({ length: rows }, (_, i) => [i, `record-${i}`, i % 7 === 0, i / 10]);
const bytes = XLSX.write(
  XLSX.utils.book_new(XLSX.utils.aoa_to_sheet(values), 'Imported benchmark'),
  { type: 'buffer', bookType: 'xlsx', compression: true },
);
const report = {
  startedAt: new Date().toISOString(),
  url: option('url', 'http://127.0.0.1:5186/'),
  rows,
  columns: 4,
  rounds,
  profile,
  inputBytes: bytes.length,
  inputSha256: createHash('sha256').update(bytes).digest('hex'),
  source: sourceFingerprint(),
  localBuild: localBuildFingerprint(),
  host: {
    node: process.version,
    platform: platform(),
    arch: arch(),
    cpu: cpus()[0].model,
    logicalCpus: cpus().length,
    memoryBytes: totalmem(),
  },
  scope:
    'Production page file input through real file/storage Workers. Timing starts at change event and ends after imported save-state mutation, grid.ready() and two animation frames. Full-cell checks, edit/undo and reopen are outside import timing. Fresh browser context each round. Frame callbacks use performance.now(); long tasks include only overlap with the measurement window. Samples cover main document only, not Worker/process memory. No universal threshold; power mode unrecorded.',
  results: [],
};
writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
const save = () => writeFileSync(output, JSON.stringify(report, null, 2) + '\n');
try {
  const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
  report.servedArtifacts = [];
  for (const file of [
    'index.html',
    ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
  ]) {
    const response = await fetch(new URL(file, report.url));
    if (!response.ok) throw new Error(`Production artifact unavailable: ${file}`);
    const actual = hash(new Uint8Array(await response.arrayBuffer()));
    if (actual !== hash(readFileSync(`apps/playground/dist/${file}`)))
      throw new Error(`Served build differs from local build: ${file}`);
    report.servedArtifacts.push({ file, sha256: actual });
  }
  save();
  for (const engine of engines) {
    const browser = await launchers[engine].launch();
    try {
      for (let round = 1; round <= rounds; round++) {
        const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
        try {
          const page = await context.newPage(),
            errors = [];
          page.on('pageerror', (error) => errors.push(String(error)));
          await page.goto(report.url);
          await page.waitForFunction(
            () =>
              !!window.opensheet?.selection.get() &&
              !document.querySelector('#scene-select').disabled,
          );
          await page.evaluate(async (profile) => {
            await window.opensheet.getGrid().ready();
            if (profile) {
              window.__importProfile = {};
              const originalPost = Worker.prototype.postMessage;
              Worker.prototype.postMessage = function (data, ...args) {
                if (data?.buffer instanceof ArrayBuffer && data.name === 'benchmark.xlsx') {
                  const sent = performance.now();
                  this.addEventListener(
                    'message',
                    (event) => {
                      window.__importProfile.fileWorkerRoundTripMs = performance.now() - sent;
                      window.__importProfile.fileWorker = event.data.profile ?? null;
                    },
                    { once: true },
                  );
                  data = { ...data, profile: true };
                }
                if (
                  data?.method === 'open' &&
                  data.args?.[1]?.sheets?.some((sheet) => sheet.blocks?.length)
                ) {
                  const sent = performance.now();
                  const observation = { phases: {} };
                  window.__importProfile.storageOpen = observation;
                  const listener = (event) => {
                    const message = event.data,
                      elapsedMs = performance.now() - sent;
                    if (message.requestId === data.id && message.progress) {
                      const progress = message.progress;
                      const phase = (observation.phases[progress.phase] ??= {
                        firstMs: elapsedMs,
                        events: 0,
                      });
                      phase.lastMs = elapsedMs;
                      phase.events++;
                      phase.completed = progress.completed;
                      phase.total = progress.total;
                      if (progress.elapsedMs !== undefined) phase.elapsedMs = progress.elapsedMs;
                    }
                    if (message.id === data.id) {
                      observation.roundTripMs = elapsedMs;
                      this.removeEventListener('message', listener);
                    }
                  };
                  this.addEventListener('message', listener);
                }
                return originalPost.call(this, data, ...args);
              };
              const app = window.opensheet,
                originalLoad = app.load.bind(app);
              app.load = async (...args) => {
                const start = performance.now();
                try {
                  return await originalLoad(...args);
                } finally {
                  window.__importProfile.workbookLoadMs = performance.now() - start;
                }
              };
            }
            let start, previous, frame, observer;
            const frames = [],
              longTasks = [];
            const tick = () => {
              const now = performance.now();
              frames.push(now - previous);
              previous = now;
              frame = requestAnimationFrame(tick);
            };
            const supportsLongTasks = PerformanceObserver.supportedEntryTypes.includes('longtask');
            document.querySelector('#file').addEventListener(
              'change',
              () => {
                start = previous = performance.now();
                frame = requestAnimationFrame(tick);
                if (supportsLongTasks) {
                  observer = new PerformanceObserver((list) =>
                    longTasks.push(
                      ...list
                        .getEntries()
                        .map((e) => ({ start: e.startTime, duration: e.duration })),
                    ),
                  );
                  observer.observe({ type: 'longtask' });
                }
              },
              { capture: true, once: true },
            );
            const stateObserver = new MutationObserver(async () => {
              if (
                start === undefined ||
                !document.querySelector('#save-state').textContent.includes('Imported locally')
              )
                return;
              stateObserver.disconnect();
              await window.opensheet.getGrid().ready();
              await new Promise((resolve) =>
                requestAnimationFrame(() => requestAnimationFrame(resolve)),
              );
              const end = performance.now(),
                durationMs = end - start;
              cancelAnimationFrame(frame);
              if (observer) {
                longTasks.push(
                  ...observer
                    .takeRecords()
                    .map((e) => ({ start: e.startTime, duration: e.duration })),
                );
                observer.disconnect();
              }
              const sorted = [...frames].sort((a, b) => a - b);
              window.__importMeasurement = {
                durationMs,
                startMs: start,
                endMs: end,
                frameSamples: frames,
                frameP95Ms: sorted[Math.floor(sorted.length * 0.95)] ?? null,
                frameMaxMs: sorted.at(-1) ?? null,
                longTasks: supportsLongTasks
                  ? longTasks
                      .filter((task) => task.start < end && task.start + task.duration > start)
                      .map((task) => ({
                        ...task,
                        overlapMs:
                          Math.min(end, task.start + task.duration) - Math.max(start, task.start),
                      }))
                  : null,
                mainHeapBytes: performance.memory?.usedJSHeapSize ?? null,
                profile: window.__importProfile ?? null,
              };
            });
            stateObserver.observe(document.querySelector('#save-state'), {
              childList: true,
              characterData: true,
              subtree: true,
            });
          }, profile);
          await page.locator('#file').setInputFiles({
            name: 'benchmark.xlsx',
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            buffer: bytes,
          });
          await page.waitForFunction(() => !!window.__importMeasurement, null, { timeout: 35000 });
          const result = await page.evaluate(async (rows) => {
            const app = window.opensheet,
              measurement = window.__importMeasurement;
            let book = app.getWorkbook(),
              sheet = book.getSheets()[0];
            const verify = async () => {
              for (let start = 0; start < rows; start += 2048) {
                const end = Math.min(start + 2048, rows);
                const actual = await sheet
                  .range({ startRow: start, endRow: end, startColumn: 0, endColumn: 4 })
                  .getValues();
                for (let i = 0; i < actual.length; i++) {
                  const row = start + i,
                    expected = [row, `record-${row}`, row % 7 === 0, row / 10];
                  if (
                    actual[i].length !== 4 ||
                    actual[i].some((value, col) => value !== expected[col])
                  )
                    throw new Error(`Cell mismatch at row ${row}`);
                }
                if (actual.length !== end - start) throw new Error('Missing imported rows');
              }
            };
            await verify();
            const editStart = performance.now();
            await sheet.range('A1').setValues([[123456]]);
            const editMs = performance.now() - editStart;
            if ((await sheet.range('A1').getValues())[0][0] !== 123456)
              throw new Error('Edit failed');
            await book.undo();
            if ((await sheet.range('A1').getValues())[0][0] !== 0) throw new Error('Undo failed');
            const id = book.id,
              reopenStart = performance.now();
            await app.open(id);
            const reopenMs = performance.now() - reopenStart;
            book = app.getWorkbook();
            sheet = book.getSheets()[0];
            await verify();
            const storage = await book.storageStats();
            if (storage.bytes > storage.budgetBytes) throw new Error('Storage budget exceeded');
            return {
              ...measurement,
              editMs,
              reopenMs,
              verifiedCells: rows * 4,
              verification: 'passed',
              storage,
            };
          }, rows);
          if (
            profile &&
            (!result.profile?.fileWorker ||
              Object.values(result.profile.fileWorker).some(
                (value) => !Number.isFinite(value) || value < 0,
              ) ||
              !Number.isFinite(result.profile.workbookLoadMs))
          )
            throw new Error('Missing or invalid import phase timings');
          if (errors.length) throw new Error(errors.join('\n'));
          report.results.push({ engine, browserVersion: browser.version(), round, ...result });
          save();
          process.stdout.write(
            JSON.stringify({
              engine,
              round,
              durationMs: result.durationMs,
              frameP95Ms: result.frameP95Ms,
              frameMaxMs: result.frameMaxMs,
              editMs: result.editMs,
              verification: result.verification,
              profile: result.profile,
            }) + '\n',
          );
        } finally {
          await context.close();
        }
      }
    } finally {
      await browser.close();
    }
  }
  if (JSON.stringify(localBuildFingerprint()) !== JSON.stringify(report.localBuild))
    throw new Error('Build changed during benchmark');
  report.finishedAt = new Date().toISOString();
  save();
} catch (error) {
  report.error = String(error);
  save();
  throw error;
}
