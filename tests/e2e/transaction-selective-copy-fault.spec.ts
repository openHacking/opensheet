import { expect, test } from '@playwright/test';

test('cancelling after selective-copy page writes preserves the committed workbook', async ({
  page,
}) => {
  test.setTimeout(60000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchCopyPages = false, copyPages = 0;
addEventListener('message', ({data}) => { if (data.watchCopyPages) watchCopyPages = true; });
const originalCopyPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalCopyPut.apply(this,args);
  if (watchCopyPages && this.name === 'records') this.transaction.addEventListener('complete',()=>{
    if (++copyPages === 2) { watchCopyPages = false; postMessage({cancelCopyAfterPages:true}); }
  },{once:true});
  return request;
};`,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 1000, columns: 4 }] }));
    let book = app.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1:A1000')
      .setValues(Array.from({ length: 1000 }, (_, i) => [i]));
    const revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    const get = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const raw = async () => {
      const db = await get(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const prefix = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          get(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            get(tx.objectStore(name).getAllKeys(prefix)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const baseline = await raw();
    const worker = (book as any).worker as Worker;
    let reached = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.cancelCopyAfterPages) {
        reached = true;
        book.cancel();
      }
    });
    worker.postMessage({ watchCopyPages: true });
    const copy = () =>
      book.transaction({}, async (tx) => {
        const sheet = tx.getSheets()[0];
        await sheet.range('B1').setValues([[999]]);
        await sheet.range('A1:A1000').copyTo('C1:C1000', { mode: 'values' });
      });
    const code = await copy().then(
      () => 'ok',
      (error) => error.code,
    );
    const after = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      keys: await raw(),
      values: await book.getSheets()[0].range('A1:C1').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    const reopened = await book.getSheets()[0].range('A1:C1').getValues();
    await copy();
    const verify = async (copied: boolean) => {
      const values = await book.getSheets()[0].range('C1:C1000').getValues();
      return values.every((row, i) => row[0] === (copied ? i : null));
    };
    const copied = await verify(true);
    await book.undo();
    const undone = await verify(false);
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return {
      reached,
      code,
      revision,
      bytes,
      baseline,
      after,
      reopened,
      copied,
      undone,
      redone: await verify(true),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.reached).toBe(true);
  expect(result.code).toBe('ABORTED');
  expect(result.after).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    keys: result.baseline,
    values: [[0, null, null]],
  });
  expect(result.reopened).toEqual([[0, null, null]]);
  expect(result.copied && result.undone && result.redone).toBe(true);
});
