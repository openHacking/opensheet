import { expect, test } from '@playwright/test';
import { createSheet } from '../../packages/core/src/model.js';

for (const phase of ['before', 'after'] as const) {
  test(`populated sheet recovers after Worker termination ${phase} head commit`, async ({
    page,
  }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let sheetTermination;
addEventListener('message', ({data}) => { if (data.sheetTermination) sheetTermination = data.sheetTermination; });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalPut.apply(this, args);
  if (!sheetTermination || this.name !== 'heads') return request;
  const phase = sheetTermination;
  sheetTermination = undefined;
  if (phase === 'before') {
    const tx = this.transaction;
    request.addEventListener('success', () => {
      postMessage({sheetTerminationReached: phase});
      const keepAlive = () => {
        const next = tx.objectStore('heads').get('__sheet_keepalive__');
        next.onsuccess = keepAlive;
      };
      keepAlive();
    }, {once: true});
  } else {
    this.transaction.oncomplete = () => postMessage({sheetTerminationReached: phase});
  }
  return request;
};`,
      });
    });
    await page.goto('/#budget');
    const initial = createSheet('Seeded', 70, 2);
    initial.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } };
    initial.cells['64:0'] = { rowId: 64, columnId: 0, input: { type: 'number', value: 7 } };
    initial.cells['64:1'] = {
      rowId: 64,
      columnId: 1,
      input: { type: 'formula', expression: 'A1+A65' },
    };
    const result = await page.evaluate(
      async ({ inputJSON, phase, modulePath }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ modulePath);
        let book = await createWorkbook({ sheets: [{ name: 'Original', rows: 2, columns: 2 }] });
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([[99]]);
        const initial = JSON.parse(inputJSON);
        const id = book.id;
        const revision = book.revision;
        const get = <T>(request: IDBRequest<T>) =>
          new Promise<T>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
        const raw = async () => {
          const db = await get(indexedDB.open('opensheet'));
          try {
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
            const decoder = new TextDecoder();
            return {
              head: Array.from(head as Uint8Array),
              recordKeys,
              sizeKeys,
              ledger: (sizes as Uint8Array[]).map((entry) =>
                JSON.parse(decoder.decode(entry.subarray(1))),
              ),
              expectedLedger: recordKeys.map((key, index) => [
                key,
                (records[index] as Uint8Array).byteLength,
              ]),
              scratch,
              physicalBytes:
                (head as Uint8Array).byteLength +
                [...records, ...sizes].reduce(
                  (total, entry) => total + (entry as Uint8Array).byteLength,
                  0,
                ),
            };
          } finally {
            db.close();
          }
        };
        const before = { bytes: (await book.storageStats()).bytes, raw: await raw() };
        const worker = (book as any).worker as Worker;
        const reached = new Promise<void>((resolve) => {
          worker.addEventListener('message', ({ data }) => {
            if (data.sheetTerminationReached === phase) {
              worker.terminate();
              resolve();
            }
          });
        });
        worker.postMessage({ sheetTermination: phase });
        const pending = book.execute({ type: 'core.sheet.add', payload: { sheet: initial } }).then(
          () => 'unexpected success',
          (error: { code?: string }) => error.code ?? 'error',
        );
        await reached;
        book.dispose();
        const outcome = await pending;
        book = await openWorkbook(id);
        const recoveredRevision = book.revision;
        const recoveredSheet = book.getSheetById(initial.id);
        const read = async () => {
          const sheet = book.getSheetById(initial.id)!;
          return [await sheet.range('A1').getValues(), await sheet.range('A65:B65').getValues()];
        };
        const values = recoveredSheet ? await read() : null;
        const recovered = { bytes: (await book.storageStats()).bytes, raw: await raw() };
        if (phase === 'before') {
          await book.execute({ type: 'core.sheet.add', payload: { sheet: initial } });
        }
        const retryOrRecovered = await read();
        await book.undo();
        const undone = !book.getSheetById(initial.id);
        await book.redo();
        const redone = await read();
        await book.close();
        book = await openWorkbook(id);
        const reopened = await read();
        await book.deleteStorage();
        return {
          outcome,
          revision,
          recoveredRevision,
          values,
          before,
          recovered,
          retryOrRecovered,
          undone,
          redone,
          reopened,
        };
      },
      {
        inputJSON: JSON.stringify(initial),
        phase,
        modulePath: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      },
    );
    expect(result.outcome).not.toBe('unexpected success');
    expect(result.recoveredRevision).toBe(result.revision + (phase === 'after' ? 1 : 0));
    const expected = [[[4]], [[7, 11]]];
    expect(result.values).toEqual(phase === 'after' ? expected : null);
    expect(result.retryOrRecovered).toEqual(expected);
    expect(result.undone).toBe(true);
    expect(result.redone).toEqual(expected);
    expect(result.reopened).toEqual(expected);
    expect(result.recovered.raw.recordKeys).toEqual(result.recovered.raw.sizeKeys);
    expect(result.recovered.raw.ledger).toEqual(result.recovered.raw.expectedLedger);
    expect(result.recovered.raw.scratch).toEqual([]);
    expect(result.recovered.raw.physicalBytes).toBe(result.recovered.bytes);
    if (phase === 'before') expect(result.recovered).toEqual(result.before);
  });
}
