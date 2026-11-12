import { expect, test } from '@playwright/test';

test('cancelling paged fill after staged writes preserves and recovers all cells', async ({
  page,
}) => {
  test.setTimeout(90_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let signalFillWrites = false, fillWrites = 0;
addEventListener('message', ({data}) => { if (data.watchFillWrites) signalFillWrites = true; });
const watchFillPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = watchFillPut.apply(this, args);
  if (signalFillWrites && this.name === 'records') {
    this.transaction.addEventListener('complete', () => {
      fillWrites++;
      if (fillWrites === 2) { signalFillWrites = false; postMessage({cancelFillNow: true}); }
    }, { once: true });
  }
  return request;
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.generate(100_000);
    let sheet = book.getSheets()[0];
    const read = async () => {
      const values = [];
      for (let row = 0; row < 1000; row += 100)
        values.push(
          ...(await sheet
            .range({ startRow: row, endRow: row + 100, startColumn: 0, endColumn: 100 })
            .getValues()),
        );
      return values;
    };
    const original = await read();
    const verify = async (filled: boolean) => {
      const values = await read();
      for (let r = 0; r < 1000; r++)
        for (let c = 0; c < 100; c++)
          if (values[r][c] !== original[filled ? 0 : r][c]) throw new Error(`Mismatch ${r},${c}`);
    };
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const get = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const rawKeys = async () => {
      const db = await get(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all(
          ['records', 'sizes', 'scratch'].map((name) =>
            get(tx.objectStore(name).getAllKeys(range)),
          ),
        );
      } finally {
        db.close();
      }
    };
    const baselineKeys = await rawKeys();
    const worker = (book as any).worker as Worker;
    let reached = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.cancelFillNow) {
        reached = true;
        book.cancel();
      }
    });
    worker.postMessage({ watchFillWrites: true });
    const code = await sheet
      .range('A1:CV1000')
      .fillDown()
      .then(
        () => 'success',
        (error) => error.code,
      );
    const failedRevision = book.revision;
    const failedBytes = (await book.storageStats()).bytes;
    const failedKeys = await rawKeys();
    await verify(false);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedRevision = book.revision;
    await verify(false);
    await sheet.range('A1:CV1000').fillDown();
    await verify(true);
    await book.undo();
    await verify(false);
    return {
      reached,
      code,
      revision,
      failedRevision,
      reopenedRevision,
      bytes,
      failedBytes,
      baselineKeys,
      failedKeys,
    };
  });
  expect(result.reached).toBe(true);
  expect(result.code).toBe('ABORTED');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.failedKeys).toEqual(result.baselineKeys);
});

for (const fault of ['page', 'head'] as const) {
  test(`paged fill survives Worker termination after durable ${fault}`, async ({ page }) => {
    test.setTimeout(90_000);
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let fillTermination, fillPageCount = 0;
addEventListener('message', ({data}) => { if (data.fillTermination) fillTermination = data.fillTermination; });
const terminateFillPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = terminateFillPut.apply(this, args);
  if (fillTermination === 'page' && this.name === 'records') {
    fillPageCount++;
    if (fillPageCount === 2) {
      fillTermination = undefined;
      this.transaction.oncomplete = () => postMessage({fillTerminationReached: 'page'});
    }
  } else if (fillTermination === 'head' && this.name === 'heads') {
    fillTermination = undefined;
    this.transaction.oncomplete = () => postMessage({fillTerminationReached: 'head'});
  }
  return request;
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (fault) => {
      const app = window.opensheet;
      let book = app.getWorkbook();
      await book.generate(100_000);
      let sheet = book.getSheets()[0];
      const readAll = async () => {
        const values = [];
        for (let startRow = 0; startRow < 1000; startRow += 100)
          values.push(
            ...(await sheet
              .range({ startRow, endRow: startRow + 100, startColumn: 0, endColumn: 100 })
              .getValues()),
          );
        return values;
      };
      const originalAll = await readAll();
      const verifyAll = async (filled: boolean, phase: string) => {
        const values = await readAll();
        for (let row = 0; row < 1000; row++)
          for (let column = 0; column < 100; column++)
            if (values[row][column] !== originalAll[filled ? 0 : row][column])
              throw new Error(`${phase} mismatch at ${row},${column}`);
      };
      const read = async () =>
        Promise.all(['A1', 'A500', 'A1000'].map((address) => sheet.range(address).getValues()));
      const original = await read();
      const revision = book.revision;
      const bytes = (await book.storageStats()).bytes;
      (book as any).options.operationTimeoutMs = 5000;
      const worker = (book as any).worker as Worker;
      let reached = '';
      worker.addEventListener('message', ({ data }) => {
        if (data.fillTerminationReached) {
          reached = data.fillTerminationReached;
          worker.terminate();
        }
      });
      worker.postMessage({ fillTermination: fault });
      const code = await sheet
        .range('A1:CV1000')
        .fillDown()
        .then(
          () => 'success',
          (error) => error.code,
        );
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      const recoveredRevision = book.revision;
      const recovered = await read();
      const recoveredBytes = (await book.storageStats()).bytes;
      await verifyAll(fault === 'head', 'recovered');
      const get = <T>(request: IDBRequest<T>) =>
        new Promise<T>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
      const db = await get(indexedDB.open('opensheet'));
      const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
      const keys = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
      const [head, recordKeys, recordValues, sizeKeys, sizeValues, scratch] = await Promise.all([
        get(tx.objectStore('heads').get(book.id)),
        get(tx.objectStore('records').getAllKeys(keys)),
        get(tx.objectStore('records').getAll(keys)),
        get(tx.objectStore('sizes').getAllKeys(keys)),
        get(tx.objectStore('sizes').getAll(keys)),
        get(tx.objectStore('scratch').getAllKeys(keys)),
      ]);
      db.close();
      const decoder = new TextDecoder();
      const ledger = sizeValues.map((entry: Uint8Array) =>
        JSON.parse(decoder.decode(entry.subarray(1))),
      );
      const expectedLedger = recordKeys.map((key, index) => [key, recordValues[index].byteLength]);
      const physicalBytes =
        head.byteLength +
        [...recordValues, ...sizeValues].reduce((sum, entry) => sum + entry.byteLength, 0);
      if (fault === 'page') await sheet.range('A1:CV1000').fillDown();
      const filled = await read();
      await verifyAll(true, 'filled');
      await book.undo();
      const undone = await read();
      await verifyAll(false, 'undo');
      await book.redo();
      const redone = await read();
      await verifyAll(true, 'redo');
      return {
        reached,
        code,
        revision,
        recoveredRevision,
        bytes,
        recoveredBytes,
        original,
        recovered,
        filled,
        undone,
        redone,
        storage: { recordKeys, sizeKeys, ledger, expectedLedger, scratch, physicalBytes },
      };
    }, fault);
    expect(result.reached).toBe(fault);
    expect(result.code).toBe('WORKER_TIMEOUT');
    expect(result.recoveredRevision).toBe(result.revision + (fault === 'head' ? 1 : 0));
    if (fault === 'page') {
      expect(result.recovered).toEqual(result.original);
      expect(result.recoveredBytes).toBe(result.bytes);
    } else
      expect(result.recovered).toEqual([
        result.original[0],
        result.original[0],
        result.original[0],
      ]);
    expect(result.filled).toEqual([result.original[0], result.original[0], result.original[0]]);
    expect(result.undone).toEqual(result.original);
    expect(result.redone).toEqual(result.filled);
    expect(result.storage.recordKeys).toEqual(result.storage.sizeKeys);
    expect(result.storage.ledger).toEqual(result.storage.expectedLedger);
    expect(result.storage.scratch).toEqual([]);
    expect(result.storage.physicalBytes).toBe(result.recoveredBytes);
  });
}

test('late quota failure reclaims staged fill pages before retry', async ({ page }) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchFillPages = false, committedFillPages = 0;
addEventListener('message', ({data}) => { if (data.failThirdFillPage) watchFillPages = true; });
const originalFillPagePut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (!watchFillPages || this.name !== 'records') return originalFillPagePut.apply(this, args);
  if (committedFillPages === 2) {
    watchFillPages = false;
    postMessage({fillPagesBeforeQuota: committedFillPages});
    throw new DOMException('Injected quota after two fill pages', 'QuotaExceededError');
  }
  this.transaction.addEventListener('complete', () => { committedFillPages++; }, { once: true });
  return originalFillPagePut.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.generate(100_000);
    let sheet = book.getSheets()[0];
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
    const original = await sheet.range('A1:A3').getValues();
    const revision = book.revision;
    let committed = 0;
    const worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.fillPagesBeforeQuota) committed = data.fillPagesBeforeQuota;
    });
    worker.postMessage({ failThirdFillPage: true });
    const code = await sheet
      .range('A1:CV1000')
      .fillDown()
      .then(
        () => 'success',
        (error) => error.code,
      );
    const failedRevision = book.revision;
    const failed = await raw();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1:A3').getValues();
    await sheet.range('A1:CV1000').fillDown();
    const filled = await sheet.range('A1:A3').getValues();
    await book.undo();
    const undone = await sheet.range('A1:A3').getValues();
    return {
      baseline,
      failed,
      committed,
      code,
      revision,
      failedRevision,
      original,
      reopened,
      filled,
      undone,
    };
  });
  expect(result.committed).toBe(2);
  expect(result.code).toBe('QuotaExceededError');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.failed).toEqual(result.baseline);
  expect(result.reopened).toEqual(result.original);
  expect(result.filled).toEqual([result.original[0], result.original[0], result.original[0]]);
  expect(result.undone).toEqual(result.original);
});

test('large paged fill preserves the old revision when head publication aborts', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let abortFillHead = false;
addEventListener('message', e => { if (e.data.abortFillHead) abortFillHead = true; });
const originalFillPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (abortFillHead && this.name === 'heads') {
    abortFillHead = false;
    this.transaction.abort();
    throw new DOMException('Injected fill head abort', 'AbortError');
  }
  return originalFillPut.apply(this, args);
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.generate(100_000);
    let sheet = book.getSheets()[0];
    const original = await sheet.range('A1:A3').getValues();
    const revision = book.revision;
    (book as any).worker.postMessage({ abortFillHead: true });
    const code = await sheet
      .range('A1:CV1000')
      .fillDown()
      .then(
        () => 'success',
        (error) => error.code,
      );
    const afterFailure = book.revision;
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopenedRevision = book.revision;
    const reopened = await sheet.range('A1:A3').getValues();
    await sheet.range('A1:CV1000').fillDown();
    const filled = await sheet.range('A1:A3').getValues();
    await book.undo();
    const undone = await sheet.range('A1:A3').getValues();
    return { code, revision, afterFailure, reopenedRevision, original, reopened, filled, undone };
  });
  expect(result.code).toBe('AbortError');
  expect(result.afterFailure).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.reopened).toEqual(result.original);
  expect(result.filled).toEqual([result.original[0], result.original[0], result.original[0]]);
  expect(result.undone).toEqual(result.original);
});
