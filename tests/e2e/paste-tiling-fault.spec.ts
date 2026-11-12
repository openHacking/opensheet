import { expect, test } from '@playwright/test';

test('cancelling a 100,000-cell tiled paste preserves and recovers all cells', async ({ page }) => {
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
  const result = await page.evaluate(async (url) => {
    const { pasteClipboard } = await import(/* @vite-ignore */ url);
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
          if (values[r][c] !== (filled ? 42 : original[r][c]))
            throw new Error(`Mismatch ${r},${c}`);
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
    const paste = () =>
      pasteClipboard(
        book,
        { sheetId: sheet.id, startRow: 0, endRow: 1000, startColumn: 0, endColumn: 100 },
        { text: '42' },
      );
    const code = await paste().then(
      () => 'success',
      (error: { code: string }) => error.code,
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
    await paste();
    await verify(true);
    await book.undo();
    await verify(false);
    await book.redo();
    await verify(true);
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify(true);
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
  }, `/@fs${process.cwd()}/packages/renderer/src/clipboard.ts`);
  expect(result.reached).toBe(true);
  expect(result.code).toBe('ABORTED');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.reopenedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.failedKeys).toEqual(result.baselineKeys);
});

test('a late oversized relocated formula rolls back previously staged paste pages', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchTile = false;
addEventListener('message', ({data}) => { if (data.watchTile) watchTile = true; });
const originalTilePut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalTilePut.apply(this, args);
  if (watchTile && this.name === 'records') this.transaction.addEventListener('complete', () => {
    watchTile = false; postMessage({ tilePageCommitted: true });
  }, { once: true });
  return request;
};`,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Tile', rows: 130, columns: 2 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[17]]);
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const request = <T>(r: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    const raw = async () => {
      const db = await request(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const prefix = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          request(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            request(tx.objectStore(name).getAllKeys(prefix)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const baseline = await raw();
    const worker = (book as any).worker as Worker;
    let staged = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.tilePageCommitted) staged = true;
    });
    worker.postMessage({ watchTile: true });
    const expression = `A1&"${'x'.repeat(32760)}"`;
    const code = await book
      .execute({
        type: 'core.cells.pasteTile',
        payload: {
          sheetId: sheet.id,
          range: { startRow: 0, endRow: 101, startColumn: 0, endColumn: 1 },
          rows: 1,
          columns: 1,
          cells: [{ input: { type: 'formula', expression }, source: { row: 0, column: 0 } }],
        },
      })
      .then(
        () => 'success',
        (error) => error.code,
      );
    const after = await raw();
    const current = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      values: await sheet.range('A1:A2').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      staged,
      code,
      baseline,
      after,
      revision,
      bytes,
      current,
      reopened: await sheet.range('A1:A2').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.staged).toBe(true);
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.after).toEqual(result.baseline);
  expect(result.current).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    values: [[17], [null]],
  });
  expect(result.reopened).toEqual([[17], [null]]);
});
