import { expect, test } from '@playwright/test';

for (const phase of ['before', 'after'] as const) {
  test(`workbook deletion is atomic when its Worker terminates ${phase} commit`, async ({
    page,
  }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let deletionCrashPhase, suppressDeletionReply = false;
const deletionCrashPost = self.postMessage.bind(self);
self.postMessage = function(message, ...rest) {
  if (suppressDeletionReply && message && message.id !== undefined) return;
  return deletionCrashPost(message, ...rest);
};
addEventListener('message', ({data}) => { if (data.deletionCrashPhase) deletionCrashPhase = data.deletionCrashPhase; });
const deletionCrashDelete = IDBObjectStore.prototype.delete;
IDBObjectStore.prototype.delete = function(...args) {
  const request = deletionCrashDelete.apply(this, args);
  if (deletionCrashPhase && this.name === 'scratch') {
    const phase = deletionCrashPhase;
    deletionCrashPhase = undefined;
    suppressDeletionReply = true;
    const tx = this.transaction;
    if (phase === 'before') request.addEventListener('success', () => {
      postMessage({ deletionCrashReached: phase });
      const keepAlive = () => { const next = tx.objectStore('heads').get('__deletion_test_keepalive__'); next.onsuccess = keepAlive; };
      keepAlive();
    });
    else tx.addEventListener('complete', () => postMessage({ deletionCrashReached: phase }));
  }
  return request;
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (phase) => {
      const source = window.opensheet.getWorkbook();
      const snapshot = await source.toJSON();
      const Workbook = source.constructor as new (input: typeof snapshot) => typeof source;
      let book = new Workbook({ ...snapshot, workbookId: `delete_crash_${crypto.randomUUID()}` });
      await book.ready();
      const id = book.id;
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      await sheet.range('A1').setValues([[84]]);
      const get = <T>(req: IDBRequest<T>) =>
        new Promise<T>((resolve, reject) => {
          req.onsuccess = () => resolve(req.result);
          req.onerror = () => reject(req.error);
        });
      const db = await get(indexedDB.open('opensheet'));
      const scratch = db.transaction('scratch', 'readwrite');
      scratch.objectStore('scratch').put(new Uint8Array([1, 2, 3]), `${id}/crash`);
      await new Promise<void>((resolve, reject) => {
        scratch.oncomplete = () => resolve();
        scratch.onabort = () => reject(scratch.error);
      });
      const capture = async () => {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
        const requests = [
          tx.objectStore('heads').get(id),
          ...['records', 'sizes', 'scratch'].flatMap((name) => [
            tx.objectStore(name).getAllKeys(range),
            tx.objectStore(name).getAll(range),
          ]),
        ];
        return JSON.stringify(await Promise.all(requests.map(get)), (_key, value) =>
          value instanceof Uint8Array ? Array.from(value) : value,
        );
      };
      const before = await capture();
      const worker = (book as any).worker as Worker;
      const reached = new Promise<void>((resolve) =>
        worker.addEventListener('message', ({ data }) => {
          if (data.deletionCrashReached === phase) {
            worker.terminate();
            resolve();
          }
        }),
      );
      worker.postMessage({ deletionCrashPhase: phase });
      void book.deleteStorage().catch(() => {});
      await reached;
      book.dispose();
      const after = await capture();
      db.close();
      if (phase === 'after')
        return { unchanged: before === after, deleted: after === '[null,[],[],[],[],[],[]]' };
      await window.opensheet.open(id);
      book = window.opensheet.getWorkbook();
      const recovered = await book.getSheets()[0].range('A1').getValues();
      const undo = await book.undo();
      const undone = await book.getSheets()[0].range('A1').getValues();
      const redo = await book.redo();
      const redone = await book.getSheets()[0].range('A1').getValues();
      await book.deleteStorage();
      return { unchanged: before === after, recovered, undo, undone, redo, redone };
    }, phase);
    if (phase === 'before')
      expect(result).toEqual({
        unchanged: true,
        recovered: [[84]],
        undo: true,
        undone: [[42]],
        redo: true,
        redone: [[84]],
      });
    else expect(result).toEqual({ unchanged: false, deleted: true });
  });
}

for (const store of ['sizes', 'scratch'] as const) {
  test(`aborted workbook deletion preserves every store when ${store} deletion fails`, async ({
    page,
  }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let failStore;
addEventListener('message', e => { if (e.data.failDeletion) failStore = e.data.failDeletion; });
const originalDelete = IDBObjectStore.prototype.delete;
IDBObjectStore.prototype.delete = function(...args) {
  if (failStore && this.name === failStore) {
    failStore = undefined;
    throw new DOMException('Injected deletion failure', 'UnknownError');
  }
  return originalDelete.apply(this, args);
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (store) => {
      const original = window.opensheet.getWorkbook();
      const snapshot = await original.toJSON();
      const Workbook = original.constructor as new (input: typeof snapshot) => typeof original;
      const book = new Workbook({ ...snapshot, workbookId: `delete_${crypto.randomUUID()}` });
      await book.ready();
      let disposalNotifications = 0;
      book.onDispose(() => disposalNotifications++);
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      await sheet.range('A1').setValues([[84]]);
      const revision = book.revision;
      const request = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const scratch = db.transaction('scratch', 'readwrite');
      scratch.objectStore('scratch').put(new Uint8Array([0, 1, 2]), `${book.id}/test`);
      await new Promise<void>((resolve, reject) => {
        scratch.oncomplete = () => resolve();
        scratch.onabort = () => reject(scratch.error);
      });
      const capture = async () => {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        const reqs = [
          tx.objectStore('heads').get(book.id),
          ...['records', 'sizes', 'scratch'].flatMap((name) => [
            tx.objectStore(name).getAllKeys(range),
            tx.objectStore(name).getAll(range),
          ]),
        ];
        const values = await Promise.all(
          reqs.map(
            (req) =>
              new Promise<any>((resolve, reject) => {
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
              }),
          ),
        );
        return JSON.stringify(values, (_key, value) =>
          value instanceof Uint8Array ? Array.from(value) : value,
        );
      };
      const before = await capture();
      (book as any).worker.postMessage({ failDeletion: store });
      let failed = false;
      try {
        await book.deleteStorage();
      } catch {
        failed = true;
      }
      const after = await capture();
      const usable = await sheet.range('A1').getValues();
      const notificationsAfterFailure = disposalNotifications;
      db.close();
      book.dispose();
      await window.opensheet.open(book.id);
      const recovered = window.opensheet.getWorkbook();
      const value = await recovered.getSheets()[0].range('A1').getValues();
      const undo = await recovered.undo();
      const undone = await recovered.getSheets()[0].range('A1').getValues();
      const redo = await recovered.redo();
      const redone = await recovered.getSheets()[0].range('A1').getValues();
      const result = {
        failed,
        notificationsAfterFailure,
        disposalNotifications,
        usable,
        unchanged: before === after,
        revision: recovered.revision,
        expectedRevision: revision + 2,
        value,
        undo,
        undone,
        redo,
        redone,
      };
      await recovered.deleteStorage();
      return result;
    }, store);
    expect(result).toEqual({
      failed: true,
      notificationsAfterFailure: 0,
      disposalNotifications: 1,
      usable: [[84]],
      unchanged: true,
      revision: result.expectedRevision,
      expectedRevision: result.expectedRevision,
      value: [[84]],
      undo: true,
      undone: [[42]],
      redo: true,
      redone: [[84]],
    });
  });
}

test('deletion prevents concurrent operations and closes the instance after commit', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const source = window.opensheet.getWorkbook();
    const snapshot = await source.toJSON();
    const Workbook = source.constructor as new (input: typeof snapshot) => typeof source;
    const book = new Workbook({ ...snapshot, workbookId: `delete_${crypto.randomUUID()}` });
    await book.ready();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[42]]);
    const worker = (book as any).worker as Worker;
    const send = worker.postMessage.bind(worker);
    let held: any;
    worker.postMessage = ((message: any) => {
      if (message.method === 'clear') held = message;
      else send(message);
    }) as typeof worker.postMessage;
    const deleting = book.deleteStorage();
    let pendingCode = '';
    try {
      await sheet.range('A1').setValues([[99]]);
    } catch (error) {
      pendingCode = (error as any).code;
    }
    let duplicateCode = '';
    try {
      await book.deleteStorage();
    } catch (error) {
      duplicateCode = (error as any).code;
    }
    send(held);
    await deleting;
    let afterCode = '';
    try {
      await sheet.range('A1').setValues([[100]]);
    } catch (error) {
      afterCode = (error as any).code;
    }
    const request = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
    const range = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
    const reqs = [
      tx.objectStore('heads').get(book.id),
      ...['records', 'sizes', 'scratch'].map((name) => tx.objectStore(name).count(range)),
    ];
    const values = await Promise.all(
      reqs.map(
        (req) =>
          new Promise<any>((resolve, reject) => {
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
          }),
      ),
    );
    db.close();
    return {
      pendingCode,
      duplicateCode,
      afterCode,
      hasHead: values[0] !== undefined,
      counts: values.slice(1),
      other: await source.getSheets()[0].range('F11').getValues(),
    };
  });
  expect(result).toEqual({
    pendingCode: 'BUSY',
    duplicateCode: 'BUSY',
    afterCode: 'DISPOSED',
    hasHead: false,
    counts: [0, 0, 0],
    other: [[13097]],
  });
});
