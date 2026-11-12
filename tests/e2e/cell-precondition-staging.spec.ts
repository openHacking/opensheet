import { expect, test } from '@playwright/test';

test('a late original-input conflict rolls back every staged fill page', async ({ page }) => {
  test.setTimeout(90_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchPreconditionPages = false, preconditionPages = 0;
addEventListener('message', ({data}) => { if (data.watchPreconditionPages) watchPreconditionPages = true; });
const preconditionPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = preconditionPut.apply(this, args);
  if (watchPreconditionPages && this.name === 'records') {
    this.transaction.addEventListener('complete', () => {
      preconditionPages++;
      if (preconditionPages === 2) {
        watchPreconditionPages = false;
        postMessage({preconditionPagesReached: preconditionPages});
      }
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
    await sheet.range('A1').setValues([[12345]]);
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
    const verify = async (filled: boolean, lastOverride = false) => {
      const values = await read();
      for (let row = 0; row < 1000; row++)
        for (let column = 0; column < 100; column++) {
          const expected =
            lastOverride && row === 999 && column === 0 ? 999 : original[filled ? 0 : row][column];
          if (values[row][column] !== expected) throw new Error(`Mismatch ${row},${column}`);
        }
    };
    const metadata = book.sheetData(sheet.id);
    const expected = {
      rowId: metadata.rowOrder[999],
      columnId: metadata.columnOrder[0],
      input: (await book.getCell(sheet.id, 999, 0))!.input,
    };
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
        const [head, ...stores] = await Promise.all([
          get(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            get(tx.objectStore(name).getAllKeys(keys)),
          ),
        ]);
        return { head: Array.from(head as Uint8Array), stores };
      } finally {
        db.close();
      }
    };
    const before = await raw();
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    let pages = 0;
    const worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.preconditionPagesReached) pages = data.preconditionPagesReached;
    });
    worker.postMessage({ watchPreconditionPages: true });
    const execute = (input: typeof expected.input) =>
      book.transaction({}, async (tx) => {
        await tx.getSheets()[0].range('A1:CV1000').fillDown();
        await tx.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [
              {
                row: 999,
                column: 0,
                input: { type: 'number', value: 999 },
                expected: { ...expected, input },
              },
            ],
          },
        });
      });
    const code = await execute(expected.input).then(
      () => 'success',
      (error: { code?: string }) => error.code,
    );
    const after = await raw();
    const failedRevision = book.revision;
    const failedBytes = (await book.storageStats()).bytes;
    await verify(false);
    await book.undo();
    const undoneAnchor = await sheet.range('A1').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify(false);
    await execute({ type: 'number', value: 12345 });
    await verify(true, true);
    await book.undo();
    await verify(false);
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    await verify(true, true);
    return {
      pages,
      code,
      before,
      after,
      revision,
      failedRevision,
      bytes,
      failedBytes,
      undoneAnchor,
    };
  });
  expect(result.pages).toBe(2);
  expect(result.code).toBe('REVISION_CONFLICT');
  expect(result.failedRevision).toBe(result.revision);
  expect(result.failedBytes).toBe(result.bytes);
  expect(result.after).toEqual(result.before);
  expect(result.undoneAnchor).toEqual([[0]]);
});
