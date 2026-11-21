import { expect, test } from '@playwright/test';

test('rejects a malformed main-thread IndexedDB before allocating resources and retries', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const id = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[731]]);
    return book.id;
  });
  await page.addInitScript(() => {
    const scope = window as any;
    scope.savedIndexedDBDescriptor = Object.getOwnPropertyDescriptor(window, 'indexedDB');
    scope.workerStarts = 0;
    scope.channelStarts = 0;
    const WorkerClass = window.Worker;
    const ChannelClass = window.BroadcastChannel;
    window.Worker = class extends WorkerClass {
      constructor(...args: ConstructorParameters<typeof Worker>) {
        super(...args);
        scope.workerStarts++;
      }
    };
    window.BroadcastChannel = class extends ChannelClass {
      constructor(...args: ConstructorParameters<typeof BroadcastChannel>) {
        super(...args);
        scope.channelStarts++;
      }
    };
    Object.defineProperty(window, 'indexedDB', { configurable: true, value: { open: null } });
  });
  await page.reload();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('STORAGE_UNAVAILABLE');
  expect(
    await page.evaluate(() => ({
      workers: (window as any).workerStarts,
      channels: (window as any).channelStarts,
      reference: localStorage.getItem('opensheet.scene.budget'),
    })),
  ).toEqual({ workers: 0, channels: 0, reference: id });
  await page.evaluate(() =>
    Object.defineProperty(window, 'indexedDB', (window as any).savedIndexedDBDescriptor),
  );
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[731]]);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
});

test('rejects malformed Worker-side IndexedDB and reopens unchanged data after retry', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const id = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[947]]);
    return book.id;
  });
  let inject = true;
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    const body = await response.text();
    await route.fulfill({
      response,
      body: inject
        ? `${body}\nObject.defineProperty(globalThis, 'indexedDB', { configurable: true, value: { open: null } });`
        : body,
    });
  });
  await page.reload();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('STORAGE_UNAVAILABLE');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  inject = false;
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[947]]);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
});
