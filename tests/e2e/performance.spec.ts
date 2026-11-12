import { test, expect } from '@playwright/test';
import { readForPoll } from './read-for-poll.js';
test('performance example measures a complete workload and restores stored data', async ({
  page,
}) => {
  test.setTimeout(180000);
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.locator('#perf-run').click();
  await expect.poll(() => page.evaluate(() => window.opensheet.element.inert)).toBe(true);
  await expect(page.locator('#perf-progress')).toContainText('Finished', { timeout: 150000 });
  expect(await page.evaluate(() => window.opensheet.element.inert)).toBe(false);
  await expect(page.locator('#perf-results tr')).toHaveCount(1);
  const download = page.waitForEvent('download');
  await page.locator('#perf-download').click();
  const file = await (await download).path();
  const fs = await import('node:fs');
  const result = JSON.parse(fs.readFileSync(file!, 'utf8'));
  expect(result.protocol.editRowSampling).toBe('evenly-spaced-inclusive-v1');
  expect(result.protocol.scrollSampling).toBe('evenly-spaced-2d-permuted-inclusive-v2');
  expect(result.protocol.viewportValidation).toBe('viewport-rectangle-cached-value-display-v2');
  fs.writeFileSync(
    `test-results/performance-${test.info().project.name}.json`,
    JSON.stringify(result, null, 2),
  );
  expect(result.results[0].cells).toBe(10000);
  expect(result.results[0].rounds.length).toBeGreaterThan(0);
  for (const round of result.results[0].rounds) {
    expect(round.editSamplesMs).toHaveLength(100);
    expect(round.commitSamplesMs).toHaveLength(100);
    expect(round.viewportSamplesMs).toHaveLength(100);
    for (let i = 0; i < 100; i++)
      expect(round.commitSamplesMs[i] + round.viewportSamplesMs[i]).toBeCloseTo(
        round.editSamplesMs[i],
        6,
      );
    expect([...round.commitSamplesMs].sort((a: number, b: number) => a - b)[94]).toBe(
      round.commitP95Ms,
    );
    expect([...round.viewportSamplesMs].sort((a: number, b: number) => a - b)[94]).toBe(
      round.viewportP95Ms,
    );
    expect(round.scrollSamplesMs).toHaveLength(300);
    expect(round.viewportChecks).toBeGreaterThan(1200);
    expect(round.scrollExtents.observedTop[0]).toBeLessThanOrEqual(1);
    expect(round.scrollExtents.observedLeft[0]).toBeLessThanOrEqual(1);
    expect(round.scrollExtents.observedTop[1]).toBeGreaterThanOrEqual(
      round.scrollExtents.maxTop - 1,
    );
    expect(round.scrollExtents.observedLeft[1]).toBeGreaterThanOrEqual(
      round.scrollExtents.maxLeft - 1,
    );
    expect([...round.editSamplesMs].sort((a: number, b: number) => a - b)[94]).toBe(
      round.editP95Ms,
    );
    expect([...round.scrollSamplesMs].sort((a: number, b: number) => a - b)[284]).toBe(
      round.scrollP95Ms,
    );
  }
  if (result.results[0].passed) expect(result.results[0].rounds).toHaveLength(3);
  const before = await page.evaluate(
    async () => await window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
  );
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
    ),
  ).toEqual(before);
  expect(errors).toEqual([]);
});
test('stopping capacity generation preserves a recoverable workbook', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.locator('#perf-cells').fill('1000000');
  await page.locator('#perf-run').click();
  await page.locator('#perf-stop').click();
  await expect(page.locator('#perf-progress')).toContainText('Stopped');
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await expect(page.locator('.os-grid')).toBeVisible();
});
test('multiple tabs share revisions and only one writer can commit', async ({ page, context }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const id = await page.evaluate(() => window.opensheet.getWorkbook().id),
    second = await context.newPage();
  await second.goto('/#budget');
  await expect(second.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await second.evaluate(() => window.opensheet.getWorkbook().isReadOnly)).toBe(true);
  await page.evaluate(
    async () =>
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1')
        .setValues([['shared']]),
  );
  await expect.poll(() => readForPoll(second, 'A1')).toEqual([['shared']]);
  expect(await second.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
  await second.close();
});

test('worker interruption recovers the previous head and reclaims uncommitted blocks', async ({
  page,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['stable']]);
    const id = book.id,
      revision = book.revision;
    const internal = book as unknown as {
      worker: Worker;
      options: { operationTimeoutMs: number; onProgress: (p: { phase: string }) => void };
    };
    internal.options.operationTimeoutMs = 500;
    internal.options.onProgress = (p) => {
      if (p.phase === 'generate') internal.worker.terminate();
    };
    let code = '';
    try {
      await book.generate(1000000);
    } catch (error) {
      code = (error as { code: string }).code;
    }
    await app.open(id, { database: 'opensheet-performance' });
    const restored = app.getWorkbook();
    return {
      code,
      revision: restored.revision,
      expectedRevision: revision,
      values: await restored.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result).toEqual({
    code: 'WORKER_TIMEOUT',
    revision: result.expectedRevision,
    expectedRevision: result.expectedRevision,
    values: [['stable']],
  });
  await page.locator('#perf-clear').click();
  await expect(page.locator('#perf-progress')).toContainText('Test data cleared');
  expect(
    await page.evaluate(
      async () => await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
});
test('native quota failures preserve committed values in a real storage worker', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let injectQuota = false;
addEventListener('message', e => { if(e.data.injectQuota) injectQuota = true; });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) { if(injectQuota && this.name === 'records') { injectQuota = false; throw new DOMException('Injected disk quota failure', 'QuotaExceededError'); } return originalPut.apply(this,args); };`,
    });
  });
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['stable']]);
    const revision = book.revision;
    (book as unknown as { worker: Worker }).worker.postMessage({ injectQuota: true });
    let code = '';
    try {
      await book.generate(10000);
    } catch (error) {
      code = (error as { code: string }).code;
    }
    return {
      code,
      unchanged: book.revision === revision,
      values: await book.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result).toEqual({ code: 'QuotaExceededError', unchanged: true, values: [['stable']] });
});

test('rapid scrolling and cache eviction reload the latest viewport without stale data', async ({
  page,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook();
    await book.generate(300000);
    await app.attachWorkbook(book);
    const grid = app.getGrid()!,
      scroller = grid.scroller,
      sheet = book.getSheets()[0];
    for (let i = 0; i < 24; i++) {
      scroller.scrollTop = i * 2400;
      await grid.ready();
    }
    scroller.scrollTop = 60000;
    scroller.dispatchEvent(new Event('scroll'));
    scroller.scrollTop = 0;
    await sheet.range('A1').setValues([['latest']]);
    await grid.ready();
    await sheet.setFreeze(sheet.rowCount, 100);
    await grid.ready();
    return {
      value: book.peekValue(sheet.id, 0, 0),
      loaded: book.isLoaded(sheet.id, 0, 0),
      stats: await book.storageStats(),
    };
  });
  expect(result.value).toBe('latest');
  expect(result.loaded).toBe(true);
  expect(result.stats.cacheBytes).toBeLessThanOrEqual(64 * 1048576);
  expect(result.stats.viewportCacheBytes).toBeLessThanOrEqual(8 * 1048576);
});

test('leaving the performance scene cancels work before attaching the next scene', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.locator('#perf-cells').fill('1000000');
  await page.locator('#perf-run').click();
  await page.evaluate(() => {
    location.hash = 'sales';
  });
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  await expect(page.getByRole('tab', { name: 'Sales Dashboard', exact: true })).toBeVisible();
  expect(errors).toEqual([]);
});

test('performance acceptance rejects a redo that reports success without restoring the edit', async ({
  page,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() => {
    const prototype = Object.getPrototypeOf(window.opensheet.getWorkbook());
    prototype.redo = async () => true;
  });
  await page.locator('#perf-run').click();
  await expect(page.locator('#perf-progress')).toContainText('Finished', { timeout: 150000 });
  const download = page.waitForEvent('download');
  await page.locator('#perf-download').click();
  const file = await (await download).path();
  const fs = await import('node:fs');
  const result = JSON.parse(fs.readFileSync(file!, 'utf8'));
  expect(result.results[0]).toMatchObject({
    cells: 10000,
    passed: false,
    failure: 'Redo validation failed',
  });
});

test('performance acceptance rejects a viewport whose cached values are unavailable', async ({
  page,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() => {
    const prototype = Object.getPrototypeOf(window.opensheet.getWorkbook());
    const isLoaded = prototype.isLoaded;
    prototype.isLoaded = function (sheet: string, row: number, column: number) {
      return row === 2 && column === 2 ? false : isLoaded.call(this, sheet, row, column);
    };
  });
  await page.locator('#perf-run').click();
  await expect(page.locator('#perf-progress')).toContainText('Finished', { timeout: 150000 });
  const download = page.waitForEvent('download');
  await page.locator('#perf-download').click();
  const file = await (await download).path();
  const fs = await import('node:fs');
  const result = JSON.parse(fs.readFileSync(file!, 'utf8'));
  expect(result.results[0].passed).toBe(false);
  expect(result.results[0].failure).toBe('Viewport validation failed at row 2, column 2');
});

test('opt-in Worker timing reports one bounded sample per measured edit', async ({ page }) => {
  test.setTimeout(180000);
  await page.goto('/?trace-ops=1#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.locator('#perf-run').click();
  await expect(page.locator('#perf-progress')).toContainText('Finished', { timeout: 150000 });
  const download = page.waitForEvent('download');
  await page.locator('#perf-download').click();
  const file = await (await download).path();
  const fs = await import('node:fs');
  const result = JSON.parse(fs.readFileSync(file!, 'utf8'));
  expect(result.protocol.operationTiming).toBe(true);
  expect(result.results[0].rounds.length).toBeGreaterThan(0);
  for (const round of result.results[0].rounds) {
    expect(round.workerQueueSamplesMs).toHaveLength(100);
    expect(round.scrollReadTraces).toHaveLength(300);
    expect(round.scrollPhaseTraces).toHaveLength(300);
    for (let i = 0; i < 300; i++) {
      const phase = round.scrollPhaseTraces[i];
      const durations = [
        phase.layoutMs,
        phase.previousFetchMs,
        phase.viewportFetchMs,
        phase.frameWaitMs,
        phase.drawMs,
      ];
      for (const duration of durations) expect(duration).toBeGreaterThanOrEqual(0);
      expect(durations.reduce((sum, duration) => sum + duration, 0)).toBeLessThanOrEqual(
        round.scrollSamplesMs[i] + 1,
      );
    }
    for (const trace of round.scrollReadTraces) {
      expect(trace.count).toBeGreaterThanOrEqual(0);
      expect(trace.maxRoundTripMs).toBeGreaterThanOrEqual(0);
      expect(trace.sumExecutionMs).toBeGreaterThanOrEqual(0);
      expect(trace.maxExecutionMs).toBeGreaterThanOrEqual(0);
      expect(trace.maxQueueMs).toBeGreaterThanOrEqual(0);
      expect(trace.maxOutsideWorkerMs).toBeGreaterThanOrEqual(0);
      expect(trace.sumOutsideWorkerMs).toBeGreaterThanOrEqual(trace.maxOutsideWorkerMs);
      expect(trace.pageCacheHits).toBeGreaterThanOrEqual(0);
      expect(trace.pageCacheMisses).toBeGreaterThanOrEqual(0);
      expect(trace.absentPages).toBeGreaterThanOrEqual(0);
      expect(trace.storageMs).toBeGreaterThanOrEqual(0);
      expect(trace.decodeMs).toBeGreaterThanOrEqual(0);
      expect(trace.inflateMs).toBeGreaterThanOrEqual(0);
      expect(trace.pageDecodeMs).toBeGreaterThanOrEqual(0);
      expect(trace.inflateMs + trace.pageDecodeMs).toBeCloseTo(trace.decodeMs, 5);
      if (trace.count > 0)
        expect(trace.pageCacheHits + trace.pageCacheMisses + trace.absentPages).toBeGreaterThan(0);
      expect(trace.storageMs + trace.decodeMs).toBeLessThanOrEqual(trace.sumExecutionMs + 1);
    }
    expect(round.workerExecutionSamplesMs).toHaveLength(100);
    expect(round.rpcSamplesMs).toHaveLength(100);
    expect(round.workerActiveSamples).toHaveLength(100);
    expect(round.workerQueuedAheadSamples).toHaveLength(100);
    for (let i = 0; i < 100; i++) {
      expect(round.workerQueueSamplesMs[i]).toBeGreaterThanOrEqual(0);
      expect(round.workerExecutionSamplesMs[i]).toBeGreaterThanOrEqual(0);
      expect(round.rpcSamplesMs[i]).toBeGreaterThanOrEqual(0);
      expect(round.rpcSamplesMs[i]).toBeLessThanOrEqual(round.commitSamplesMs[i] + 1);
    }
  }
});
