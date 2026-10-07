import { spawn } from 'node:child_process';
import { readFileSync, writeFileSync } from 'node:fs';
import { platform, arch } from 'node:os';
import { chromium } from '@playwright/test';
import { createSnapshot, key } from '../packages/core/dist/index.js';

const sizes = (
  process.argv.find((arg) => arg.startsWith('--sizes='))?.slice(8) ?? '10000,100000,200000'
)
  .split(',')
  .map(Number);
const baselinePath = new URL('../benchmarks/browser-baseline.json', import.meta.url);
const median = (values) => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)];
const percentile = (values, p) =>
  [...values].sort((a, b) => a - b)[Math.ceil(values.length * p) - 1];
const fixture = (count) => {
  const rows = Math.ceil(count / 20);
  const snapshot = createSnapshot({ sheets: [{ name: 'Performance', rows, columns: 24 }] });
  const sheet = snapshot.sheets[0];
  sheet.freeze = { rows: 1, columns: 1 };
  sheet.rows[sheet.rowOrder[0]] = { size: 36 };
  for (let n = 0; n < count; n++) {
    const row = Math.floor(n / 20),
      column = n % 20;
    const rowId = sheet.rowOrder[row],
      columnId = sheet.columnOrder[column];
    sheet.cells[key(rowId, columnId)] = {
      rowId,
      columnId,
      input: { type: 'number', value: n },
    };
  }
  for (let row = 0; row + 1 < Math.min(rows, 2000); row += 2)
    sheet.merges.push({ startRow: row, endRow: row + 2, startColumn: 20, endColumn: 22 });
  return snapshot;
};

let server;
let browser;
try {
  let ready = false;
  try {
    ready = (await fetch('http://127.0.0.1:5173/')).ok;
  } catch {}
  if (!ready) server = spawn('pnpm', ['dev'], { stdio: 'ignore', detached: true });
  for (let i = 0; i < 80; i++) {
    if (server && server.exitCode !== null) throw new Error('Vite exited before becoming ready');
    try {
      ready = (await fetch('http://127.0.0.1:5173/')).ok;
    } catch {}
    if (ready) break;
    await new Promise((resolve) => setTimeout(resolve, 250));
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
      '20 populated columns, 4 empty columns, frozen first row and column, up to 1000 nonoverlapping merges',
    scenarios: {},
  };
  for (const count of sizes) {
    if (![10000, 100000, 200000].includes(count)) throw new Error(`Unsupported size: ${count}`);
    const page = await browser.newPage({ viewport: results.viewport });
    await page.goto('http://127.0.0.1:5173/');
    await page.locator('.os-grid').waitFor();
    const snapshot = fixture(count);
    const measurement = await page.evaluate(async (data) => {
      const app = window.opensheet;
      const longTasks = [];
      const measurementStart = performance.now();
      const observer =
        'PerformanceObserver' in window
          ? new PerformanceObserver((list) => {
              for (const entry of list.getEntries())
                if (entry.startTime >= measurementStart) longTasks.push(entry.duration);
            })
          : null;
      try {
        observer?.observe({ type: 'longtask' });
      } catch {}
      const frame = () =>
        new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const loadMs = [];
      for (let i = 0; i < 3; i++) {
        const start = performance.now();
        await app.load(data);
        await frame();
        loadMs.push(performance.now() - start);
      }
      const book = app.getWorkbook(),
        sheet = book.getSheets()[0],
        sheetId = sheet.id;
      const editMs = [];
      for (let i = 0; i < 5; i++) {
        const t = performance.now();
        sheet.range('A2').setValues([[i]]);
        await frame();
        editMs.push(performance.now() - t);
      }
      const selectMs = [];
      for (let i = 0; i < 20; i++) {
        const t = performance.now();
        const row = 10 + ((i * 37) % (data.sheets[0].rowOrder.length - 10));
        app.selection.set({
          sheetId,
          startRow: row,
          endRow: row + 1,
          startColumn: i % 20,
          endColumn: (i % 20) + 1,
        });
        await frame();
        selectMs.push(performance.now() - t);
      }
      const scroller = document.querySelector('.os-scroll');
      const scrollMs = [];
      for (let i = 0; i < 30; i++) {
        const t = performance.now();
        scroller.scrollTop = (i * 977) % scroller.scrollHeight;
        scroller.scrollLeft = i % 2 ? 2400 : 0;
        await frame();
        scrollMs.push(performance.now() - t);
      }
      observer?.disconnect();
      return {
        loadMs,
        editMs,
        selectMs,
        scrollMs,
        longTasks,
        serializedBytes: JSON.stringify(data).length,
      };
    }, snapshot);
    const session = await page.context().newCDPSession(page);
    await session.send('Performance.enable');
    const metrics = await session.send('Performance.getMetrics');
    const heapBytes = metrics.metrics.find((metric) => metric.name === 'JSHeapUsedSize')?.value;
    await session.send('HeapProfiler.collectGarbage');
    const retainedMetrics = await session.send('Performance.getMetrics');
    const retainedBytes = retainedMetrics.metrics.find(
      (metric) => metric.name === 'JSHeapUsedSize',
    )?.value;
    results.scenarios[count] = {
      loadP50Ms: Math.round(median(measurement.loadMs)),
      loadP95Ms: Math.round(percentile(measurement.loadMs, 0.95)),
      editP50Ms: Math.round(median(measurement.editMs)),
      editP95Ms: Math.round(percentile(measurement.editMs, 0.95)),
      selectP50Ms: Math.round(median(measurement.selectMs)),
      selectP95Ms: Math.round(percentile(measurement.selectMs, 0.95)),
      scrollP50Ms: Math.round(median(measurement.scrollMs)),
      scrollP95Ms: Math.round(percentile(measurement.scrollMs, 0.95)),
      longTaskCount: measurement.longTasks.length,
      longTaskP95Ms: measurement.longTasks.length
        ? Math.round(percentile(measurement.longTasks, 0.95))
        : 0,
      heapBeforeGcMiB: heapBytes ? Math.round(heapBytes / 1024 / 1024) : null,
      retainedHeapMiB: retainedBytes ? Math.round(retainedBytes / 1024 / 1024) : null,
      serializedMiB: Math.round(measurement.serializedBytes / 1024 / 1024),
    };
    await page.close();
    console.error(`Measured ${count} cells`);
  }
  if (process.argv.includes('--compare')) {
    const baseline = JSON.parse(readFileSync(baselinePath, 'utf8'));
    if (
      baseline.browser === results.browser &&
      baseline.platform === results.platform &&
      baseline.arch === results.arch
    ) {
      for (const count of sizes)
        for (const metric of ['loadP95Ms', 'editP95Ms', 'selectP95Ms', 'scrollP95Ms']) {
          const old = baseline.scenarios[count]?.[metric],
            current = results.scenarios[count][metric];
          if (old && current > Math.max(old * 1.2, old + 10))
            throw new Error(`${count} ${metric} regressed: ${old} → ${current} ms`);
        }
    } else console.error('Environment differs from baseline; comparison skipped.');
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
