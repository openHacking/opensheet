import { expect, test } from '@playwright/test';

test('a read-only tab cannot clear the performance workbook', async ({ page, context }) => {
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
  await second.locator('#perf-clear').click();
  await expect(second.locator('#perf-progress')).toContainText(/read only|another tab/i);
  expect(
    await second.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  await expect(second.locator('.os-grid')).toBeVisible();
  await second.locator('#perf-run').click();
  await expect(second.locator('#perf-progress')).toContainText('read only');
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
});

test('an aborted clear retains the prior workbook and permits retry', async ({ page }) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let abortPerformanceClear = false;
addEventListener('message', event => { if (event.data.abortPerformanceClear) abortPerformanceClear = true; });
const priorClearPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (abortPerformanceClear && this.name === 'heads') {
    abortPerformanceClear = false;
    this.transaction.abort();
    throw new DOMException('Injected clear head abort', 'AbortError');
  }
  return priorClearPut.apply(this, args);
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
    (book as any).worker.postMessage({ abortPerformanceClear: true });
    return book.revision;
  });
  await page.locator('#perf-clear').click();
  await expect(page.locator('#perf-progress')).toContainText('Injected clear head abort');
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
  await page.locator('#perf-clear').click();
  await expect(page.locator('#perf-progress')).toContainText('Test data cleared');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[null]]);
});

test('clear preserves an active formula draft', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]),
  );
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  await formula.fill('unsaved draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  await page.evaluate(() => document.getElementById('perf-clear')!.click());
  await expect(page.locator('#perf-progress')).toContainText('pending changes');
  await expect(formula).toHaveValue('unsaved draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
});

test('capacity testing preserves an active draft and its saved workbook', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]),
  );
  const formula = page.getByRole('combobox', { name: 'Formula bar' });
  await formula.fill('unsaved draft');
  await page.evaluate(() => document.getElementById('perf-run')!.click());
  await expect(page.locator('#perf-progress')).toContainText('pending changes');
  await expect(formula).toHaveValue('unsaved draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
});

test('capacity testing rejects a write during the storage estimate', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const initial = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    Object.defineProperty(navigator, 'storage', {
      configurable: true,
      value: {
        estimate: () =>
          new Promise<StorageEstimate>((resolve) => {
            (window as any).resumeCapacityEstimate = () =>
              resolve({ usage: 0, quota: 1_000_000_000_000 });
          }),
      },
    });
    return book.id;
  });
  await page.locator('#perf-run').click();
  await expect.poll(() => page.evaluate(() => !!(window as any).resumeCapacityEstimate)).toBe(true);
  await expect(page.locator('#perf-run')).toBeDisabled();
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[4321]]);
    (window as any).resumeCapacityEstimate();
  });
  await expect(page.locator('#perf-progress')).toContainText(
    'Workbook changed during test preparation',
  );
  expect(
    await page.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    })),
  ).toEqual({ id: initial, value: [[4321]] });
  await expect(page.locator('#perf-run')).toBeEnabled();
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[4321]]);
});

test('leaving during a storage estimate does not start a capacity test', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
    Object.defineProperty(navigator, 'storage', {
      configurable: true,
      value: {
        estimate: () =>
          new Promise<StorageEstimate>((resolve) => {
            (window as any).resumeCapacityEstimate = () =>
              resolve({ usage: 0, quota: 1_000_000_000_000 });
          }),
      },
    });
  });
  await page.locator('#perf-run').click();
  await expect.poll(() => page.evaluate(() => !!(window as any).resumeCapacityEstimate)).toBe(true);
  await page.evaluate(() => {
    location.hash = 'budget';
  });
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    delete (navigator as any).storage;
    (window as any).resumeCapacityEstimate();
  });
  await page.evaluate(() => {
    location.hash = 'performance';
  });
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
});

test('a stalled storage estimate can be stopped without changing the workbook', async ({
  page,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
    Object.defineProperty(navigator, 'storage', {
      configurable: true,
      value: {
        estimate: () =>
          new Promise<StorageEstimate>((resolve) => {
            (window as any).resumeCapacityEstimate = () =>
              resolve({ usage: 0, quota: 1_000_000_000_000 });
          }),
      },
    });
  });
  await page.locator('#perf-run').click();
  await expect.poll(() => page.evaluate(() => !!(window as any).resumeCapacityEstimate)).toBe(true);
  await expect(page.locator('#perf-stop')).toBeEnabled();
  await page.locator('#perf-stop').click();
  await expect(page.locator('#perf-progress')).toContainText('Stopped');
  await expect(page.locator('#perf-run')).toBeEnabled();
  const summary = await page.locator('#perf-summary').textContent();
  await page.evaluate(async () => {
    (window as any).resumeCapacityEstimate();
    await new Promise((resolve) => setTimeout(resolve, 0));
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
  });
  expect(await page.locator('#perf-summary').textContent()).toBe(summary);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
});

test('leaving performance does not cancel an unrelated pending edit', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    let pendingId = -1;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute' && pendingId === -1) pendingId = message.id;
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (event.data.id === pendingId) (window as any).pendingPerformanceEdit = event;
      else handler.call(worker, event);
    };
    (window as any).releasePerformanceEdit = () => {
      handler.call(worker, (window as any).pendingPerformanceEdit);
      worker.postMessage = post;
      worker.onmessage = handler;
    };
    (window as any).performanceEdit = book
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
  });
  await expect.poll(() => page.evaluate(() => !!(window as any).pendingPerformanceEdit)).toBe(true);
  await page.evaluate(() => {
    location.hash = 'budget';
  });
  await expect(page.locator('#notice')).toContainText('Wait for saving');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(async () => {
    (window as any).releasePerformanceEdit();
    await (window as any).performanceEdit;
  });
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[421]]);
});

test('leaving waits for a pending clear before attaching another scene', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() => {
    const app = window.opensheet;
    const original = app.loadJSON.bind(app);
    app.loadJSON = async (...args) => {
      await new Promise<void>((resolve) => ((window as any).resumePendingClear = resolve));
      await original(...args);
    };
  });
  await page.locator('#perf-clear').click();
  await expect.poll(() => page.evaluate(() => !!(window as any).resumePendingClear)).toBe(true);
  await page.evaluate(async () => {
    const changed = new Promise<void>((resolve) =>
      addEventListener('hashchange', () => resolve(), { once: true }),
    );
    location.hash = 'budget';
    await changed;
  });
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(() => (window as any).resumePendingClear());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].name)).toBe(
    'Launch budget',
  );
});

test('a failure after clear publication reports an uncertain result', async ({ page }) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(async () => {
    const app = window.opensheet;
    await app
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[421]]);
    const original = app.loadJSON.bind(app);
    app.loadJSON = async (...args) => {
      await original(...args);
      throw new Error('Injected post-clear failure');
    };
  });
  await page.locator('#perf-clear').click();
  await expect(page.locator('#perf-progress')).toContainText('Clear may have completed');
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
