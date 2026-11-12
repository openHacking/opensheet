import { expect, test } from '@playwright/test';

test('readers invalidate cached data only after a confirmed deletion', async ({
  page,
  context,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let failDelete = false;
addEventListener('message', e => { if (e.data.failDelete) failDelete = true; });
const remove = IDBObjectStore.prototype.delete;
IDBObjectStore.prototype.delete = function(...args) {
  if (failDelete && this.name === 'scratch') {
    failDelete = false;
    throw new DOMException('Injected deletion failure', 'UnknownError');
  }
  return remove.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[321]]);
  });
  const reader = await context.newPage();
  await reader.goto('/#budget');
  await expect(reader.locator('#take-write-lock')).toBeVisible();
  const loaded = await reader.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await book.prefetch(sheet.id, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    return book.peekValue(sheet.id, 0, 0);
  });
  expect(loaded).toBe(321);
  const spurious = await reader.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const worker = (book as any).worker as Worker;
    const send = worker.postMessage.bind(worker);
    let target = -1;
    const checked = new Promise<void>((resolve) => {
      worker.addEventListener('message', (event) => {
        if (event.data.id === target) resolve();
      });
    });
    worker.postMessage = ((message: any) => {
      if (message.method === 'headExists') target = message.id;
      send(message);
    }) as typeof worker.postMessage;
    const channel = new BroadcastChannel(`opensheet:opensheet:${book.id}`);
    channel.postMessage({ deleted: true });
    await checked;
    worker.postMessage = send as typeof worker.postMessage;
    channel.close();
    return book.persistenceState.state;
  });
  expect(spurious).toBe('read-only');
  const failed = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    (book as any).worker.postMessage({ failDelete: true });
    try {
      await book.deleteStorage();
      return false;
    } catch {
      return true;
    }
  });
  expect(failed).toBe(true);
  expect(await reader.evaluate(() => window.opensheet.getWorkbook().persistenceState.state)).toBe(
    'read-only',
  );
  expect(
    await reader.evaluate(
      async () =>
        (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
    ),
  ).toBe(321);
  await page.evaluate(async () => window.opensheet.getWorkbook().deleteStorage());
  await expect
    .poll(() => reader.evaluate(() => window.opensheet.getWorkbook().persistenceState.error?.code))
    .toBe('WORKBOOK_DELETED');
  const after = await reader.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    let error = '';
    try {
      await sheet.range('A1').getValues();
    } catch (cause) {
      error = (cause as any).code;
    }
    return { cached: book.peekValue(sheet.id, 0, 0), error };
  });
  expect(after).toEqual({ cached: { error: '#LOADING' }, error: 'WORKBOOK_DELETED' });
  await reader.close();
});

test('a failed ordinary refresh clears previously cached data', async ({ page, context }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const reader = await context.newPage();
  await reader.goto('/#budget');
  await expect(reader.locator('#take-write-lock')).toBeVisible();
  await reader.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await book.prefetch(sheet.id, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  });
  await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const request = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').delete(book.id);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    ((book as any).channel as BroadcastChannel).postMessage({ revision: book.revision });
  });
  await expect
    .poll(() => reader.evaluate(() => window.opensheet.getWorkbook().persistenceState.error?.code))
    .toBe('CORRUPT_STORAGE');
  expect(
    await reader.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      return book.peekValue(book.getSheets()[0].id, 0, 0);
    }),
  ).toEqual({ error: '#LOADING' });
  await reader.close();
});
