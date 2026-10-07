import { test, expect } from '@playwright/test';
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
  fs.writeFileSync(
    `test-results/performance-${test.info().project.name}.json`,
    JSON.stringify(result, null, 2),
  );
  expect(result.results[0].cells).toBe(10000);
  expect(result.results[0].rounds.length).toBeGreaterThan(0);
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
  await expect
    .poll(() =>
      second.evaluate(
        async () =>
          (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
      ),
    )
    .toBe('shared');
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
    await app.open(id, { database: 'opensheet-performance-v4' });
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
