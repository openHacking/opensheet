import { expect, test } from '@playwright/test';

test('sort reclaims multiple committed output pages after a late quota failure', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchOutput = false, committed = 0;
addEventListener('message', ({data}) => { if (data.failThirdSortPage) watchOutput = true; });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (!watchOutput || this.name !== 'records') return originalPut.apply(this, args);
  if (committed === 2) {
    watchOutput = false;
    postMessage({sortPagesBeforeQuota: committed});
    throw new DOMException('Injected quota after two sort pages', 'QuotaExceededError');
  }
  this.transaction.addEventListener('complete', () => { committed++; }, { once: true });
  return originalPut.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.insertRows(100, 30);
    const range = sheet.range('A1:A130');
    await range.setValues(Array.from({ length: 130 }, (_, i) => [130 - i]));
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
    const baseline = await raw();
    const revision = book.revision;
    let committed = 0;
    const worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.sortPagesBeforeQuota) committed = data.sortPagesBeforeQuota;
    });
    worker.postMessage({ failThirdSortPage: true });
    const code = await sheet.sort('A1:A130', 0).then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    const failed = await raw();
    const failedRevision = book.revision;
    const failedValues = await range.getValues();
    await sheet.sort('A1:A130', 0);
    const sorted = await range.getValues();
    await book.undo();
    return {
      code,
      committed,
      baseline,
      failed,
      revision,
      failedRevision,
      failedValues,
      sorted,
      undone: await range.getValues(),
    };
  });
  const descending = Array.from({ length: 130 }, (_, i) => [130 - i]);
  expect(result.committed).toBe(2);
  expect(result.code).toBe('QuotaExceededError');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.failed).toEqual(result.baseline);
  expect(result.failedValues).toEqual(descending);
  expect(result.sorted).toEqual([...descending].reverse());
  expect(result.undone).toEqual(descending);
});

for (const fault of [
  'abort-head',
  'after-head',
  'quota-scratch',
  'quota-records',
  'quota-sizes',
  'quota-heads',
  'terminate-scratch',
  'terminate-head',
] as const) {
  test(`sort publication is atomic at ${fault}`, async ({ page }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let sortFault, failIndex = false;
addEventListener('message', ({data}) => { if (data.sortFault) sortFault = data.sortFault; });
const put = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (sortFault === 'quota-' + this.name) {
    postMessage({sortFaultReached: sortFault});
    sortFault = undefined;
    throw new DOMException('Injected sort quota failure', 'QuotaExceededError');
  }
  const request = put.apply(this, args);
  if ((sortFault === 'terminate-scratch' && this.name === 'scratch') ||
      (sortFault === 'terminate-head' && this.name === 'heads')) {
    const fault = sortFault;
    sortFault = undefined;
    // Hold the engine continuation after durable commit so termination cannot race its reply.
    this.transaction.oncomplete = () => postMessage({sortFaultReached: fault});
    return request;
  }
  if (this.name !== 'heads' || !sortFault) return request;
  const fault = sortFault;
  sortFault = undefined;
  if (fault === 'abort-head') this.transaction.abort();
  else {
    const complete = this.transaction.oncomplete;
    this.transaction.oncomplete = function(event) {
      failIndex = true;
      complete.call(this, event);
    };
  }
  postMessage({sortFaultReached: fault});
  return request;
};
const set = Map.prototype.set;
Map.prototype.set = function(...args) {
  if (failIndex) { failIndex = false; throw new Error('Injected post-head sort failure'); }
  return set.apply(this, args);
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (fault) => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setValues([[9], [1]]);
      await sheet.range('B1:B2').setFormulas([['A1*2'], ['A2*2']]);
      await sheet.range('A1').setStyle({ bold: true });
      const revision = book.revision;
      const baselineBytes = (await book.storageStats()).bytes;
      const worker = (book as any).worker as Worker;
      let reached = '';
      const terminate = fault.startsWith('terminate-');
      if (terminate) (book as any).options.operationTimeoutMs = 5000;
      worker.addEventListener('message', ({ data }) => {
        if (data.sortFaultReached) {
          reached = data.sortFaultReached;
          if (terminate) worker.terminate();
        }
      });
      worker.postMessage({ sortFault: fault });
      const outcome = await sheet.sort('A1:B2', 0).then(
        () => 'success',
        (error: { code: string }) => error.code,
      );
      const quota = fault.startsWith('quota-');
      const beforeRetry = quota
        ? {
            revision: book.revision,
            values: await sheet.range('A1:B2').getValues(),
            bytes: (await book.storageStats()).bytes,
          }
        : undefined;
      if (quota) await sheet.sort('A1:B2', 0);
      await window.opensheet.open(book.id);
      const reopened = window.opensheet.getWorkbook();
      const read = async () => {
        const sh = reopened.getSheets()[0];
        const styleRow = (await sh.range('A1').getValues())[0][0] === 9 ? 0 : 1;
        const style = reopened.getStyle((await reopened.getCell(sh.id, styleRow, 0))?.styleId);
        return { values: await sh.range('A1:B2').getValues(), bold: style?.bold };
      };
      const current = await read();
      const recoveredRevision = reopened.revision;
      const stats = await reopened.storageStats();
      const get = <T>(request: IDBRequest<T>) =>
        new Promise<T>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
      const db = await get(indexedDB.open('opensheet'));
      const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
      const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
      const records = tx.objectStore('records');
      const sizes = tx.objectStore('sizes');
      const [head, recordKeys, recordValues, sizeKeys, sizeValues, scratch] = await Promise.all([
        get(tx.objectStore('heads').get(book.id)),
        get(records.getAllKeys(range)),
        get(records.getAll(range)),
        get(sizes.getAllKeys(range)),
        get(sizes.getAll(range)),
        get(tx.objectStore('scratch').getAllKeys(range)),
      ]);
      db.close();
      const decoder = new TextDecoder();
      const ledger = sizeValues.map((entry: Uint8Array) =>
        JSON.parse(decoder.decode(entry.subarray(1))),
      );
      const expectedLedger = recordKeys.map((key, index) => [key, recordValues[index].byteLength]);
      const physicalBytes =
        head.byteLength +
        [...recordValues, ...sizeValues].reduce((total, entry) => total + entry.byteLength, 0);
      if (fault === 'abort-head' || fault === 'terminate-scratch')
        await reopened.getSheets()[0].sort('A1:B2', 0);
      const sorted = await read();
      await reopened.undo();
      const undone = await read();
      await reopened.redo();
      return {
        reached,
        outcome,
        beforeRetry,
        baselineBytes,
        revision,
        recoveredRevision,
        storage: {
          recordKeys,
          sizeKeys,
          ledger,
          expectedLedger,
          scratch,
          physicalBytes,
          chargedBytes: stats.bytes,
        },
        current,
        sorted,
        undone,
        redone: await read(),
      };
    }, fault);
    const original = {
      values: [
        [9, 18],
        [1, 2],
      ],
      bold: true,
    };
    const sorted = {
      values: [
        [1, 2],
        [9, 18],
      ],
      bold: true,
    };
    expect(result.reached).toBe(fault);
    const quota = fault.startsWith('quota-');
    const oldHead = fault === 'abort-head' || fault === 'terminate-scratch';
    expect(result.outcome).toBe(
      quota
        ? 'QuotaExceededError'
        : fault.startsWith('terminate-')
          ? 'WORKER_TIMEOUT'
          : fault === 'abort-head'
            ? 'ABORTED'
            : 'WORKER_FAILED',
    );
    if (quota)
      expect(result.beforeRetry).toEqual({
        revision: result.revision,
        values: original.values,
        bytes: result.baselineBytes,
      });
    expect(result.recoveredRevision).toBe(result.revision + (oldHead ? 0 : 1));
    expect(result.current).toEqual(oldHead ? original : sorted);
    expect(result.storage.recordKeys.length).toBeGreaterThan(0);
    expect(result.storage.sizeKeys).toEqual(result.storage.recordKeys);
    expect(result.storage.ledger).toEqual(result.storage.expectedLedger);
    expect(result.storage.scratch).toEqual([]);
    expect(result.storage.chargedBytes).toBe(result.storage.physicalBytes);
    expect(result.sorted).toEqual(sorted);
    expect(result.undone).toEqual(original);
    expect(result.redone).toEqual(sorted);
  });
}
