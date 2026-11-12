import { expect, test } from '@playwright/test';

test('preserves the original pre-commit error when staged-record cleanup also fails', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let failHead = false, failCleanup = false;
addEventListener('message', ({data}) => {
  if (data.armCleanupFault) { failHead = true; failCleanup = true; }
});
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (failHead && this.name === 'heads') {
    failHead = false;
    postMessage({cleanupFaultStage: 'head'});
    throw new DOMException('Injected head quota', 'QuotaExceededError');
  }
  return originalPut.apply(this, args);
};
const originalDelete = IDBObjectStore.prototype.delete;
IDBObjectStore.prototype.delete = function(...args) {
  if (!failHead && failCleanup && this.name === 'records') {
    failCleanup = false;
    postMessage({cleanupFaultStage: 'cleanup'});
    throw new DOMException('Injected cleanup failure', 'UnknownError');
  }
  return originalDelete.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[7]]);
    const get = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const raw = async () => {
      const db = await get(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const keys = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        const [head, records, sizes, scratch] = await Promise.all([
          get(tx.objectStore('heads').get(book.id)),
          get(tx.objectStore('records').getAllKeys(keys)),
          get(tx.objectStore('sizes').getAllKeys(keys)),
          get(tx.objectStore('scratch').getAllKeys(keys)),
        ]);
        return { head: Array.from(head as Uint8Array), records, sizes, scratch };
      } finally {
        db.close();
      }
    };
    const before = await raw();
    const beforeBytes = (await book.storageStats()).bytes;
    const revision = book.revision;
    const stages: string[] = [];
    const worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.cleanupFaultStage) stages.push(data.cleanupFaultStage);
    });
    worker.postMessage({ armCleanupFault: true });
    const error = await sheet
      .range('A1')
      .setValues([[9]])
      .then(
        () => 'success',
        (failure: { code: string }) => failure.code,
      );
    const failed = await raw();
    const failedBytes = (await book.storageStats()).bytes;
    const failedRevision = book.revision;
    const failedValue = await sheet.range('A1').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const recovered = await raw();
    const recoveredBytes = (await book.storageStats()).bytes;
    const recoveredValue = await sheet.range('A1').getValues();
    await sheet.range('A1').setValues([[9]]);
    await book.undo();
    const undone = await sheet.range('A1').getValues();
    await book.redo();
    return {
      stages,
      error,
      before,
      beforeBytes,
      failed,
      failedBytes,
      recovered,
      recoveredBytes,
      revision,
      failedRevision,
      failedValue,
      recoveredValue,
      undone,
      redone: await sheet.range('A1').getValues(),
    };
  });
  expect(result.stages).toEqual(['head', 'cleanup']);
  expect(result.error).toBe('QuotaExceededError');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.failed.head).toEqual(result.before.head);
  expect(result.failed.records.length).toBeGreaterThan(result.before.records.length);
  expect(result.failed.records).toEqual(result.failed.sizes);
  expect(result.failedBytes).toBeGreaterThan(result.beforeBytes);
  expect(result.recovered).toEqual(result.before);
  expect(result.recoveredBytes).toBe(result.beforeBytes);
  expect(result.failedValue).toEqual([[7]]);
  expect(result.recoveredValue).toEqual([[7]]);
  expect(result.undone).toEqual([[7]]);
  expect(result.redone).toEqual([[9]]);
});
