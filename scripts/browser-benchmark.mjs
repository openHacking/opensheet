import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { platform, arch } from 'node:os';
import { chromium } from '@playwright/test';
const sizes = (
  process.argv.find((arg) => arg.startsWith('--sizes='))?.slice(8) ?? '10000,100000,200000'
)
  .split(',')
  .map(Number);
const baselinePath = new URL('../benchmarks/browser-baseline.json', import.meta.url);
const percentile = (values, p) =>
  [...values].sort((a, b) => a - b)[Math.ceil(values.length * p) - 1];
let server, browser;
try {
  let ready = false;
  try {
    ready = (await fetch('http://127.0.0.1:5173/')).ok;
  } catch {}
  if (!ready) server = spawn('pnpm', ['dev'], { stdio: 'ignore', detached: true });
  for (let i = 0; i < 80 && !ready; i++) {
    if (server && server.exitCode !== null) throw new Error('Vite exited before becoming ready');
    await new Promise((resolve) => setTimeout(resolve, 250));
    try {
      ready = (await fetch('http://127.0.0.1:5173/')).ok;
    } catch {}
  }
  if (!ready) throw new Error('Vite did not become ready');
  browser = await chromium.launch();
  const results = {
    date: new Date().toISOString(),
    node: process.version,
    platform: platform(),
    arch: arch(),
    browser: browser.version(),
    headless: true,
    viewport: { width: 1280, height: 800 },
    fixture:
      'paged: 100 columns, fixed seed, 60% numbers / 30% 32-byte ASCII / 10% booleans; IndexedDB, Worker and bounded caches; scroll samples include viewport loading and painting',
    scenarios: {},
  };
  for (const count of sizes) {
    if (!Number.isSafeInteger(count) || count < 10000 || count > 10000000 || count % 10000)
      throw new Error(`Invalid size: ${count}`);
    const page = await browser.newPage({ viewport: results.viewport });
    await page.goto('http://127.0.0.1:5173/#performance');
    await page.locator('#scene-title').filter({ hasText: 'Performance Lab' }).waitFor();
    const measurement = await page.evaluate(async (count) => {
      const app = window.opensheet,
        generationStart = performance.now();
      await app.getWorkbook().generate(count);
      await app.attachWorkbook(app.getWorkbook());
      const generateMs = performance.now() - generationStart;
      const id = app.getWorkbook().id,
        loadMs = [],
        editMs = [],
        selectMs = [],
        scrollMs = [];
      for (let i = 0; i < 3; i++) {
        const start = performance.now();
        await app.open(id, { database: 'opensheet-performance' });
        loadMs.push(performance.now() - start);
      }
      const book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      for (let i = -5; i < 100; i++) {
        const start = performance.now();
        await sheet.range('A2').setValues([[i]]);
        await app.getGrid().ready();
        if (i >= 0) editMs.push(performance.now() - start);
      }
      for (let i = 0; i < 20; i++) {
        const row = 10 + ((i * 337) % (sheet.rowCount - 10)),
          start = performance.now();
        app.selection.set({
          sheetId: sheet.id,
          startRow: row,
          endRow: row + 1,
          startColumn: (i * 7) % 100,
          endColumn: ((i * 7) % 100) + 1,
        });
        await app.getGrid().ready();
        selectMs.push(performance.now() - start);
      }
      const frame = () => new Promise((resolve) => requestAnimationFrame(resolve)),
        scroller = app.getGrid().scroller;
      for (let i = -30; i < 300; i++) {
        const start = await frame();
        scroller.scrollTop =
          (Math.max(0, i) * 977) % Math.max(1, scroller.scrollHeight - scroller.clientHeight);
        scroller.scrollLeft =
          (Math.max(0, i) * 239) % Math.max(1, scroller.scrollWidth - scroller.clientWidth);
        await app.getGrid().ready();
        const end = performance.now();
        if (i >= 0) scrollMs.push(end - start);
      }
      await app.getGrid().ready();
      const storage = await book.storageStats();
      return {
        generateMs,
        loadMs,
        editMs,
        selectMs,
        scrollMs,
        storage,
        quota: await navigator.storage.estimate(),
      };
    }, count);
    const session = await page.context().newCDPSession(page);
    await session.send('Performance.enable');
    const metrics = await session.send('Performance.getMetrics');
    await session.send('HeapProfiler.collectGarbage');
    const retained = await session.send('Performance.getMetrics');
    const heap = (m) =>
      Math.round(
        (m.metrics.find((metric) => metric.name === 'JSHeapUsedSize')?.value ?? 0) / 1048576,
      );
    results.scenarios[count] = {
      generateMs: Math.round(measurement.generateMs),
      loadP50Ms: Math.round(percentile(measurement.loadMs, 0.5)),
      loadP95Ms: Math.round(percentile(measurement.loadMs, 0.95)),
      editP50Ms: Math.round(percentile(measurement.editMs, 0.5)),
      editP95Ms: Math.round(percentile(measurement.editMs, 0.95)),
      selectP95Ms: Math.round(percentile(measurement.selectMs, 0.95)),
      scrollFrameP95Ms: Number(percentile(measurement.scrollMs, 0.95).toFixed(2)),
      encodedMiB: Number((measurement.storage.bytes / 1048576).toFixed(2)),
      workerCacheReservationMiB: Number((measurement.storage.cacheBytes / 1048576).toFixed(2)),
      viewportCacheReservationMiB: Number(
        (measurement.storage.viewportCacheBytes / 1048576).toFixed(2),
      ),
      mainHeapBeforeGcMiB: heap(metrics),
      mainHeapRetainedMiB: heap(retained),
      originUsageMiB: Number(((measurement.quota.usage ?? 0) / 1048576).toFixed(2)),
    };
    await page.close();
    console.error(`Measured ${count} cells`);
  }
  if (process.argv.includes('--compare')) {
    const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
    if (
      baseline.fixture === results.fixture &&
      baseline.browser === results.browser &&
      baseline.platform === results.platform &&
      baseline.arch === results.arch
    ) {
      for (const count of sizes)
        for (const metric of ['loadP95Ms', 'editP95Ms', 'selectP95Ms', 'scrollFrameP95Ms']) {
          const old = baseline.scenarios[count]?.[metric],
            current = results.scenarios[count][metric];
          if (old && current > Math.max(old * 1.2, old + 10))
            throw new Error(`${count} ${metric} regressed: ${old} → ${current} ms`);
        }
    } else console.error('Environment or workload differs; comparison skipped.');
  }
  if (process.argv.includes('--write-baseline'))
    writeFileSync(baselinePath, JSON.stringify(results, null, 2) + '\n');
  console.log(JSON.stringify(results, null, 2));
} finally {
  await browser?.close();
  if (server?.pid) {
    try {
      process.kill(-server.pid, 'SIGTERM');
    } catch (error) {
      if (error.code !== 'ESRCH') throw error;
    }
  }
}
