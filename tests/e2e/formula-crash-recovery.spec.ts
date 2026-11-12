import { expect, test } from '@playwright/test';

for (const phase of ['before', 'after'] as const) {
  test(`overlapping formulas recover after Worker termination ${phase} head commit`, async ({
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
let formulaCrashPhase;
addEventListener('message', ({data}) => { if (data.formulaCrashPhase) formulaCrashPhase = data.formulaCrashPhase; });
const formulaCrashPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = formulaCrashPut.apply(this, args);
  if (formulaCrashPhase && this.name === 'heads') {
    const phase = formulaCrashPhase;
    formulaCrashPhase = undefined;
    const tx = this.transaction;
    if (phase === 'before') request.addEventListener('success', () => {
      const keepAlive = () => { const next = tx.objectStore('heads').get('__formula_crash_keepalive__'); next.onsuccess = keepAlive; };
      keepAlive();
      postMessage({formulaCrashReached: phase});
    });
    else tx.addEventListener('complete', () => postMessage({formulaCrashReached: phase}));
  }
  return request;
};`,
      });
    });
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(
      async ({ phase, moduleUrl }) => {
        const { createWorkbookFile } = await import(/* @vite-ignore */ moduleUrl);
        const app = window.opensheet;
        await app.load(
          createWorkbookFile({
            sheets: [
              { name: 'Source', rows: 512, columns: 2 },
              { name: 'Output', rows: 512, columns: 2 },
            ],
          }),
        );
        let book = app.getWorkbook();
        const [source, output] = book.getSheets();
        await source.range('A1:A32').setValues(Array.from({ length: 32 }, () => [1]));
        await output
          .range('A1:A512')
          .setFormulas(Array.from({ length: 512 }, (_, row) => [`SUM(Source!$A$1:$A$32)+${row}`]));
        await source.range('A1').setValues([[3]]);
        await source.range('A1').setValues([[1]]);
        const verify = async (value: number) => {
          const sheets = book.getSheets();
          const sourceValues = await sheets[0].range('A1:A32').getValues();
          if (
            sourceValues.length !== 32 ||
            sourceValues.some((entry, index) => entry[0] !== (index === 0 ? value : 1))
          )
            throw new Error('Source range mismatch');
          const values = await sheets[1].range('A1:A512').getValues();
          if (values.length !== 512) throw new Error('Incomplete formula range');
          for (let row = 0; row < values.length; row++)
            if (values[row][0] !== value + 31 + row) throw new Error(`Formula mismatch at ${row}`);
          return values.length;
        };
        const baselineCount = await verify(1);
        const revision = book.revision;
        (book as any).options.operationTimeoutMs = 3000;
        const worker = (book as any).worker as Worker;
        let reached = '';
        worker.addEventListener('message', ({ data }) => {
          if (data.formulaCrashReached) {
            reached = data.formulaCrashReached;
            worker.terminate();
          }
        });
        worker.postMessage({ formulaCrashPhase: phase });
        const code = await source
          .range('A1')
          .setValues([[2]])
          .then(
            () => 'success',
            (error: any) => error.code,
          );
        await app.open(book.id);
        book = app.getWorkbook();
        const recoveredRevision = book.revision;
        const recoveredCount = await verify(phase === 'before' ? 1 : 2);
        const chargedBytes = (await book.storageStats()).bytes;
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
        const expectedLedger = recordKeys.map((key, index) => [
          key,
          recordValues[index].byteLength,
        ]);
        const ledgerMatches =
          JSON.stringify(ledger) === JSON.stringify(expectedLedger) &&
          JSON.stringify(sizeKeys) === JSON.stringify(recordKeys);
        const bytesMatch =
          chargedBytes ===
          head.byteLength +
            [...recordValues, ...sizeValues].reduce((sum, entry) => sum + entry.byteLength, 0);
        const undo = await book.undo();
        const undoCount = await verify(phase === 'before' ? 3 : 1);
        const redo = await book.redo();
        const redoCount = await verify(phase === 'before' ? 1 : 2);
        return {
          reached,
          code,
          revision,
          recoveredRevision,
          baselineCount,
          recoveredCount,
          ledgerMatches,
          bytesMatch,
          scratchCount: scratch.length,
          undo,
          undoCount,
          redo,
          redoCount,
        };
      },
      { phase, moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts` },
    );
    expect(result).toEqual({
      reached: phase,
      code: 'WORKER_TIMEOUT',
      revision: result.revision,
      recoveredRevision: result.revision + (phase === 'after' ? 1 : 0),
      baselineCount: 512,
      recoveredCount: 512,
      ledgerMatches: true,
      bytesMatch: true,
      scratchCount: 0,
      undo: true,
      undoCount: 512,
      redo: true,
      redoCount: 512,
    });
  });
}
