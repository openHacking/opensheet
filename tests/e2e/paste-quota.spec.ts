import { expect, test } from '@playwright/test';

for (const styled of [false, true])
  for (const failureAt of [1, 2]) {
    test(`quota on record ${failureAt} rolls back ${styled ? 'styled' : 'input-only'} paste`, async ({
      page,
    }) => {
      await page.route('**/engine.worker.js*', async (route) => {
        const response = await route.fetch();
        await route.fulfill({
          response,
          body:
            (await response.text()) +
            `
let quotaCountdown = 0;
addEventListener('message', e => { if (e.data.injectPasteQuota) quotaCountdown = e.data.injectPasteQuota; });
const quotaOriginalPut = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  if (this.name === 'records' && quotaCountdown > 0 && --quotaCountdown === 0) {
    postMessage({ faultStage: 'quota' });
    throw new DOMException('Injected quota', 'QuotaExceededError');
  }
  return quotaOriginalPut.apply(this, args);
};`,
        });
      });
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      const before = await page.evaluate(
        async ({ styled, failureAt }) => {
          const book = window.opensheet.getWorkbook(),
            sheet = book.getSheets()[0];
          await sheet.range('A1').setValues([[10]]);
          await sheet.range('A70').setValues([[20]]);
          await sheet.range('A1').setStyle({ bold: true, color: '#123456' });
          await sheet.range('A1').setNote('original');
          const snapshot = await book.toJSON();
          const revision = book.revision;
          const stages: string[] = [];
          (book as any).worker.addEventListener('message', (event: MessageEvent) => {
            if (event.data.faultStage) stages.push(event.data.faultStage);
          });
          (book as any).worker.postMessage({ injectPasteQuota: failureAt });
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
                    input: { type: 'number', value: 100 },
                    note: 'replacement',
                    ...(styled ? { style: { italic: true, color: '#abcdef' } } : {}),
                  },
                  {
                    row: 69,
                    column: 0,
                    input: { type: 'number', value: 200 },
                    ...(styled ? { style: { bold: true } } : {}),
                  },
                ],
              },
            });
          } catch (error) {
            code = (error as { code?: string }).code ?? (error as Error).name;
          }
          return { id: book.id, revision, snapshot, code, stages };
        },
        { styled, failureAt },
      );
      expect(before.stages).toContain('quota');
      expect(before.code).toBe('QuotaExceededError');
      const after = await page.evaluate(async (id) => {
        await window.opensheet.open(id);
        const book = window.opensheet.getWorkbook(),
          sheet = book.getSheets()[0];
        const snapshot = await book.toJSON();
        const stats = await book.storageStats();
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
          undone = await sheet.range('A1').getNotes();
        const redo = await book.redo(),
          redone = await sheet.range('A1').getNotes();
        return {
          revision: snapshot.revision,
          snapshot,
          charged: stats.bytes,
          bytes,
          scratch: records[3],
          undo,
          undone,
          redo,
          redone,
        };
      }, before.id);
      expect(after.snapshot).toEqual(before.snapshot);
      expect(after.revision).toBe(before.revision);
      expect(after.charged).toBe(after.bytes);
      expect(after.scratch).toEqual([]);
      expect(after).toMatchObject({
        undo: true,
        undone: [[null]],
        redo: true,
        redone: [['original']],
      });
    });
  }
