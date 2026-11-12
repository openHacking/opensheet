import { expect, test } from '@playwright/test';

test('terminating a real Worker around the head write leaves one complete revision', async ({
  page,
}) => {
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let interruptHead = false;
addEventListener('message', e => { if (e.data.interruptHead) interruptHead = true; });
const originalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const result = originalPut.apply(this, args);
  if (interruptHead && this.name === 'heads') {
    interruptHead = false;
    postMessage({faultStage: 'head-put'});
  }
  return result;
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const initial = [
      (await sheet.range('A1').getValues())[0][0],
      (await sheet.range('A65').getValues())[0][0],
    ];
    const coords = [
      { row: 0, column: 0 },
      { row: 64, column: 0 },
    ];
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: coords.map((cell, index) => ({
          ...cell,
          input: { type: 'number', value: index ? 20 : 10 },
        })),
      },
    });
    const revision = book.revision;
    const worker = (book as any).worker as Worker;
    (book as any).options.operationTimeoutMs = 5000;
    let interrupted = false;
    worker.addEventListener('message', (event) => {
      if (event.data.faultStage === 'head-put') {
        interrupted = true;
        worker.terminate();
      }
    });
    worker.postMessage({ interruptHead: true });
    let code = '';
    try {
      await book.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: coords.map((cell, index) => ({
            ...cell,
            input: { type: 'number', value: index ? 202 : 101 },
          })),
        },
      });
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    return { id: book.id, revision, code, interrupted, initial };
  });
  expect(before.interrupted).toBe(true);
  expect(before.code).toBe('WORKER_TIMEOUT');
  const after = await page.evaluate(async (id) => {
    await window.opensheet.open(id);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const values = [
      (await sheet.range('A1').getValues())[0][0],
      (await sheet.range('A65').getValues())[0][0],
    ];
    const revision = book.revision;
    const stats = await book.storageStats();
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
    const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
    const requests = [
      tx.objectStore('heads').get(id),
      tx.objectStore('records').getAll(range),
      tx.objectStore('sizes').getAll(range),
      tx.objectStore('scratch').getAllKeys(range),
    ];
    const data = await Promise.all(
      requests.map(
        (request) =>
          new Promise<any>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          }),
      ),
    );
    db.close();
    const bytes =
      (data[0] as Uint8Array).byteLength +
      [...data[1], ...data[2]].reduce((sum: number, item: Uint8Array) => sum + item.byteLength, 0);
    const undo = await book.undo();
    const undone = [
      (await sheet.range('A1').getValues())[0][0],
      (await sheet.range('A65').getValues())[0][0],
    ];
    const redo = await book.redo();
    const redone = [
      (await sheet.range('A1').getValues())[0][0],
      (await sheet.range('A65').getValues())[0][0],
    ];
    return {
      values,
      revision,
      bytes,
      chargedBytes: stats.bytes,
      scratch: data[3],
      undo,
      undone,
      redo,
      redone,
    };
  }, before.id);
  expect([before.revision, before.revision + 1]).toContain(after.revision);
  expect(after.values).toEqual(after.revision === before.revision ? [10, 20] : [101, 202]);
  expect(after.chargedBytes).toBe(after.bytes);
  expect(after.scratch).toEqual([]);
  expect(after.undo).toBe(true);
  expect(after.undone).toEqual(after.revision === before.revision ? before.initial : [10, 20]);
  expect(after.redo).toBe(true);
  expect(after.redone).toEqual(after.values);
  test.info().annotations.push({
    type: 'recovered revision',
    description: after.revision === before.revision ? 'old' : 'new',
  });
});
