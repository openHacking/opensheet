import { expect, test } from '@playwright/test';

test('reset clears the performance workbook and persists the empty fixture', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const original = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
    return { id: book.id, revision: book.revision };
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.getByRole('button', { name: 'Reset demo' })).toBeEnabled();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    return {
      id: book.id,
      revision: book.revision,
      canUndo: book.canUndo,
      canRedo: book.canRedo,
      value: await book.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result.id).toBe(original.id);
  expect(result.revision).toBeGreaterThan(original.revision);
  expect(result.value).toEqual([[null]]);
  expect(result.canUndo).toBe(false);
  expect(result.canRedo).toBe(false);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
});

test('read-only performance reset leaves the writer and data intact', async ({ page, context }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]),
  );
  const second = await context.newPage();
  await second.goto('/#performance');
  await expect(second.locator('#scene-title')).toHaveText('Performance Lab');
  expect(await second.evaluate(() => window.opensheet.getWorkbook().isReadOnly)).toBe(true);
  await second.getByRole('button', { name: 'Reset demo' }).click();
  await expect(second.locator('#notice')).toContainText('read only');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  expect(
    await second.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  await second.reload();
  await expect(second.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await second.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  await second.close();
});

test('an aborted reset keeps the prior performance data and can be retried', async ({ page }) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let abortPerformanceReset = false;
addEventListener('message', event => { if (event.data.abortPerformanceReset) abortPerformanceReset = true; });
const priorPerformancePut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (abortPerformanceReset && this.name === 'heads') {
    abortPerformanceReset = false;
    this.transaction.abort();
    throw new DOMException('Injected reset head abort', 'AbortError');
  }
  return priorPerformancePut.apply(this, args);
};`,
    });
  });
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const revision = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
    (book as any).worker.postMessage({ abortPerformanceReset: true });
    return book.revision;
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.locator('#notice')).toContainText('Injected reset head abort');
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return { revision: book.revision, value: await book.getSheets()[0].range('A1').getValues() };
    }),
  ).toEqual({ revision, value: [[421]] });
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.locator('#perf-progress')).toHaveText('Test data reset.');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
});

test('reset controls stay disabled until the reset finishes', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() => {
    const app = window.opensheet;
    const load = app.loadJSON.bind(app);
    app.loadJSON = async (...args) => {
      await load(...args);
      await new Promise<void>((resolve) => ((window as any).finishPerformanceReset = resolve));
    };
  });
  const reset = page.getByRole('button', { name: 'Reset demo' });
  await reset.click();
  await expect.poll(() => page.evaluate(() => !!(window as any).finishPerformanceReset)).toBe(true);
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[42]]),
  );
  await expect(reset).toBeDisabled();
  await expect(page.locator('#perf-run')).toBeDisabled();
  await expect(page.locator('#scene-select')).toBeDisabled();
  await page.evaluate(() => (window as any).finishPerformanceReset());
  await expect(reset).toBeEnabled();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[42]]);
});

test('a failure after reset publication reports that it may have completed', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(async () => {
    const app = window.opensheet;
    await app
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
    const load = app.loadJSON.bind(app);
    app.loadJSON = async (...args) => {
      await load(...args);
      throw new Error('Injected post-publication failure');
    };
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.locator('#notice')).toContainText('Reset may have completed');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
});
