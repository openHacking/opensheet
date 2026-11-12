import { expect, test } from '@playwright/test';

for (const styled of [false, true]) {
  test(`worker interruption around paste head keeps one complete ${styled ? 'styled' : 'input-only'} revision`, async ({
    page,
  }) => {
    test.setTimeout(45_000);
    await page.route('**/engine.worker.js*', async (route) => {
      // Retry transient connection resets while fetching the Worker for fault injection.
      const response = await route.fetch({ maxRetries: 2 });
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let interruptPasteHead = false;
addEventListener('message', e => { if (e.data.interruptPasteHead) interruptPasteHead = true; });
const pasteOriginalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const result = pasteOriginalPut.apply(this, args);
  if (interruptPasteHead && this.name === 'heads') { interruptPasteHead = false; postMessage({ faultStage: 'paste-head' }); }
  return result;
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async (styled) => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[10]]);
      await sheet.range('A65').setValues([[20]]);
      await sheet.range('A1').setStyle({ bold: true, color: '#123456' });
      await sheet.range('A1').setNote('old');
      const revision = book.revision;
      const baseStyle = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
      const worker = (book as any).worker as Worker;
      (book as any).options.operationTimeoutMs = 5000;
      let interrupted = false;
      worker.addEventListener('message', (event) => {
        if (event.data.faultStage === 'paste-head') {
          interrupted = true;
          worker.terminate();
        }
      });
      worker.postMessage({ interruptPasteHead: true });
      let code = '';
      try {
        await book.execute({
          type: 'core.cells.paste',
          payload: {
            sheetId: sheet.id,
            cells: [
              {
                row: 0,
                column: 0,
                input: { type: 'number', value: 101 },
                note: 'new',
                ...(styled ? { style: { color: '#abcdef' } } : {}),
              },
              { row: 64, column: 0, input: { type: 'number', value: 202 } },
            ],
          },
        });
      } catch (error) {
        code = (error as { code?: string }).code ?? '';
      }
      return { id: book.id, revision, code, interrupted, baseStyle };
    }, styled);
    expect(before.interrupted).toBe(true);
    expect(before.code).toBe('WORKER_TIMEOUT');
    const after = await page.evaluate(async (id) => {
      await window.opensheet.open(id);
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      const inspect = async () => ({
        values: [
          (await sheet.range('A1').getValues())[0][0],
          (await sheet.range('A65').getValues())[0][0],
        ],
        note: (await sheet.range('A1').getNotes())[0][0],
        style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
      });
      const revision = book.revision,
        state = await inspect(),
        stats = await book.storageStats();
      const request = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
      const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
      const requests = [
        tx.objectStore('heads').get(id),
        tx.objectStore('records').getAll(range),
        tx.objectStore('sizes').getAll(range),
        tx.objectStore('scratch').getAllKeys(range),
      ];
      const records = await Promise.all(
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
        (records[0] as Uint8Array).byteLength +
        [...records[1], ...records[2]].reduce(
          (sum: number, value: Uint8Array) => sum + value.byteLength,
          0,
        );
      const undo = await book.undo(),
        undone = await inspect(),
        redo = await book.redo(),
        redone = await inspect();
      return {
        revision,
        state,
        charged: stats.bytes,
        bytes,
        scratch: records[3],
        undo,
        undone,
        redo,
        redone,
      };
    }, before.id);
    expect([before.revision, before.revision + 1]).toContain(after.revision);
    const committed = after.revision === before.revision + 1;
    expect(after.state).toEqual({
      values: committed ? [101, 202] : [10, 20],
      note: committed ? 'new' : 'old',
      style: committed && styled ? { color: '#abcdef' } : before.baseStyle,
    });
    expect(after.charged).toBe(after.bytes);
    expect(after.scratch).toEqual([]);
    expect(after.undo).toBe(true);
    expect(after.redo).toBe(true);
    expect(after.redone).toEqual(after.state);
    if (committed)
      expect(after.undone).toEqual({ values: [10, 20], note: 'old', style: before.baseStyle });
  });
}
