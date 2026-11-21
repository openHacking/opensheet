import { expect, test } from '@playwright/test';

for (const location of ['undo', 'redo'] as const) {
  for (const corruption of ['collision', 'oversized'] as const) {
    test(`rejects ${corruption} ${location} style IDs before storage recovery`, async ({
      page,
      context,
    }) => {
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      const fixture = await page.evaluate(
        async ({ location, corruption }) => {
          const book = window.opensheet.getWorkbook(),
            sheet = book.getSheets()[0];
          const originalName = sheet.name;
          await sheet.range('A1').setValues([[246]]);
          await sheet.range('A1').setStyle({ bold: true });
          await sheet.range('A2').setStyle({ italic: true });
          await book.renameSheet(sheet.id, 'Renamed');
          if (location === 'redo') await book.undo();
          const id = book.id;
          await book.close();
          const request = indexedDB.open('opensheet');
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
          const get = (store: string, key: string) =>
            new Promise<Uint8Array>((resolve, reject) => {
              const read = db.transaction(store).objectStore(store).get(key);
              read.onsuccess = () => resolve(read.result);
              read.onerror = () => reject(read.error);
            });
          const decode = (bytes: Uint8Array) =>
            JSON.parse(new TextDecoder().decode(bytes.subarray(1)));
          const head = decode(await get('heads', id))[1];
          const historyId = (location === 'undo' ? head.past : head.future).at(-1);
          const history = decode(await get('records', historyId))[1];
          const metadataId = location === 'undo' ? history.before : history.after;
          if (metadataId === head.meta) throw new Error('Fixture must corrupt historical metadata');
          const raw = await get('records', metadataId),
            ledger = await get('sizes', metadataId);
          const metadata = decode(raw);
          if (metadata[1].styleIds.length < 2) throw new Error('Fixture needs two styles');
          if (corruption === 'collision') {
            metadata[1].styleIds[0] = 0;
            metadata[1].styleIds[1] = '0';
          } else metadata[1].styleIds[0] = 'a'.repeat(101);
          const altered = new Uint8Array([
            0,
            ...new TextEncoder().encode(JSON.stringify(metadata)),
          ]);
          const alteredSize = new Uint8Array([
            0,
            ...new TextEncoder().encode(JSON.stringify([metadataId, altered.length])),
          ]);
          const scratchId = `${id}/corrupt-history-sentinel`;
          const tx = db.transaction(['records', 'sizes', 'scratch'], 'readwrite');
          const done = new Promise<void>((resolve, reject) => {
            tx.oncomplete = () => resolve();
            tx.onabort = () => reject(tx.error);
          });
          tx.objectStore('records').put(altered, metadataId);
          tx.objectStore('sizes').put(alteredSize, metadataId);
          tx.objectStore('scratch').put(new Uint8Array([0, 42]), scratchId);
          await done;
          db.close();
          return {
            id,
            originalName,
            metadataId,
            scratchId,
            raw: Array.from(raw),
            ledger: Array.from(ledger),
          };
        },
        { location, corruption },
      );
      const capture = async () =>
        page.evaluate(async () => {
          const request = indexedDB.open('opensheet');
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
          const names = ['heads', 'records', 'sizes', 'scratch'],
            tx = db.transaction(names);
          const result = await Promise.all(
            names.map(async (name) => {
              const store = tx.objectStore(name);
              const read = <T>(req: IDBRequest<T>) =>
                new Promise<T>((resolve, reject) => {
                  req.onsuccess = () => resolve(req.result);
                  req.onerror = () => reject(req.error);
                });
              const [keys, values] = await Promise.all([
                read(store.getAllKeys()),
                read(store.getAll()),
              ]);
              return [
                name,
                keys.map((key, index) => [key, Array.from(values[index] as Uint8Array)]),
              ];
            }),
          );
          db.close();
          return result;
        });
      const before = await capture();
      const recovery = await context.newPage();
      await recovery.goto('/#budget');
      await expect(
        recovery.getByRole('heading', { name: 'Unable to open workbook' }),
      ).toBeVisible();
      await expect(recovery.locator('#startup-recovery-message')).toContainText(
        'Invalid stored style identities',
      );
      expect(await capture()).toEqual(before);
      expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
        fixture.id,
      );
      await page.evaluate(async (fixture) => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const tx = db.transaction(['records', 'sizes'], 'readwrite');
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve();
          tx.onabort = () => reject(tx.error);
        });
        tx.objectStore('records').put(new Uint8Array(fixture.raw), fixture.metadataId);
        tx.objectStore('sizes').put(new Uint8Array(fixture.ledger), fixture.metadataId);
        await done;
        db.close();
      }, fixture);
      await recovery.getByRole('button', { name: 'Try again' }).click();
      await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
      const result = await recovery.evaluate(
        async ({ location, originalName }) => {
          const app = window.opensheet;
          let book = app.getWorkbook();
          const initial = book.getSheets()[0].name;
          if (location === 'undo') {
            await book.undo();
            await book.redo();
          } else {
            await book.redo();
            await book.undo();
          }
          await app.open(book.id);
          book = app.getWorkbook();
          const sheet = book.getSheets()[0];
          return {
            initial,
            name: sheet.name,
            originalName,
            value: await sheet.range('A1').getValues(),
            style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
          };
        },
        { location, originalName: fixture.originalName },
      );
      expect(result.initial).toBe(location === 'undo' ? 'Renamed' : fixture.originalName);
      expect(result.name).toBe(result.initial);
      expect(result.value).toEqual([[246]]);
      expect(result.style.bold).toBe(true);
      await recovery.close();
    });
  }
}
