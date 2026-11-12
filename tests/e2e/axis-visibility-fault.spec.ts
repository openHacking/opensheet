import { expect, test } from '@playwright/test';

for (const phase of ['abort', 'before', 'after'] as const) {
  test(`full-axis visibility recovers after ${phase} head publication`, async ({ page }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let axisFault;
addEventListener('message', ({data}) => { if (data.axisFault) axisFault = data.axisFault; });
const axisPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = axisPut.apply(this, args);
  if (!axisFault || this.name !== 'heads') return request;
  const fault = axisFault;
  axisFault = undefined;
  if (fault === 'abort') {
    this.transaction.abort();
    postMessage({axisFaultReached: fault});
  } else if (fault === 'before') {
    const tx = this.transaction;
    request.addEventListener('success', () => {
      postMessage({axisFaultReached: fault});
      const keepAlive = () => { const next = tx.objectStore('heads').get('__axis_keepalive__'); next.onsuccess = keepAlive; };
      keepAlive();
    });
  } else {
    this.transaction.oncomplete = () => postMessage({axisFaultReached: fault});
  }
  return request;
};`,
      });
    });
    await page.goto('/');
    const result = await page.evaluate(
      async ({ phase, modulePath }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ modulePath);
        let book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] });
        const sheet = book.getSheets()[0];
        await sheet.range('A1').setValues([[42]]);
        await sheet.setRowHeight(1, 55);
        await sheet.setRowsHidden(1, 2, true);
        await sheet.setColumnsHidden(1, 2, true);
        const revision = book.revision;
        const id = book.id;
        const worker = (book as any).worker as Worker;
        const reached = new Promise<void>((resolve) => {
          worker.addEventListener('message', ({ data }) => {
            if (data.axisFaultReached === phase) {
              if (phase !== 'abort') worker.terminate();
              resolve();
            }
          });
        });
        worker.postMessage({ axisFault: phase });
        const operation = sheet.setAllRowsHidden(false).then(
          () => 'success',
          (error: { code?: string }) => error.code ?? 'error',
        );
        await reached;
        if (phase !== 'abort') book.dispose();
        const outcome = await operation;
        if (phase === 'abort') await book.close();
        book = await openWorkbook(id);
        const read = async () => {
          const current = book.getSheets()[0];
          const data = book.sheetData(current.id);
          return {
            rows: data.rowOrder.slice(1, 3).map((key: number) => !!data.rows[key]?.hidden),
            columns: data.columnOrder.slice(1, 3).map((key: number) => !!data.columns[key]?.hidden),
            size: data.rows[data.rowOrder[1]]?.size,
            value: await current.range('A1').getValues(),
          };
        };
        const recoveredRevision = book.revision;
        const current = await read();
        const get = <T>(request: IDBRequest<T>) =>
          new Promise<T>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
        const db = await get(indexedDB.open('opensheet'));
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const keys = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
        const [head, recordKeys, records, sizeKeys, sizes, scratch] = await Promise.all([
          get(tx.objectStore('heads').get(id)),
          get(tx.objectStore('records').getAllKeys(keys)),
          get(tx.objectStore('records').getAll(keys)),
          get(tx.objectStore('sizes').getAllKeys(keys)),
          get(tx.objectStore('sizes').getAll(keys)),
          get(tx.objectStore('scratch').getAllKeys(keys)),
        ]);
        db.close();
        const decoder = new TextDecoder();
        const ledger = sizes.map((entry: Uint8Array) =>
          JSON.parse(decoder.decode(entry.subarray(1))),
        );
        const expectedLedger = recordKeys.map((key, index) => [key, records[index].byteLength]);
        const physicalBytes =
          head.byteLength +
          [...records, ...sizes].reduce((total, entry) => total + entry.byteLength, 0);
        const chargedBytes = (await book.storageStats()).bytes;
        await book.undo();
        const undone = await read();
        await book.redo();
        const redone = await read();
        await book.deleteStorage();
        return {
          outcome,
          revision,
          recoveredRevision,
          current,
          undone,
          redone,
          recordKeys,
          sizeKeys,
          ledger,
          expectedLedger,
          scratch,
          physicalBytes,
          chargedBytes,
        };
      },
      { phase, modulePath: `/@fs${process.cwd()}/packages/core/src/index.ts` },
    );
    const state = (rows: boolean[], columns: boolean[]) => ({
      rows,
      columns,
      size: 55,
      value: [[42]],
    });
    expect(result.outcome).not.toBe('success');
    expect(result.recoveredRevision).toBe(result.revision + (phase === 'after' ? 1 : 0));
    expect(result.current).toEqual(
      state(phase === 'after' ? [false, false] : [true, true], [true, true]),
    );
    expect(result.undone).toEqual(
      state([true, true], phase === 'after' ? [true, true] : [false, false]),
    );
    expect(result.redone).toEqual(result.current);
    expect(result.sizeKeys).toEqual(result.recordKeys);
    expect(result.ledger).toEqual(result.expectedLedger);
    expect(result.scratch).toEqual([]);
    expect(result.physicalBytes).toBe(result.chargedBytes);
  });
}
