import { expect, test } from '@playwright/test';

for (const direction of ['undo', 'redo'] as const) {
  for (const corruption of ['style', 'sheet'] as const) {
    test(`rejects ${direction} before publishing a bad historical ${corruption} reference`, async ({
      page,
    }) => {
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      const fixture = await page.evaluate(
        async ({ direction, corruption }) => {
          const book = window.opensheet.getWorkbook(),
            sheet = book.getSheets()[0];
          await sheet.range('A1').setValues([[372]]);
          await sheet.range('A1').setStyle({ bold: true });
          await sheet.range('A1').setStyle({ italic: true });
          if (direction === 'redo') await book.undo();
          const id = book.id;
          await book.close();
          const open = indexedDB.open('opensheet');
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            open.onsuccess = () => resolve(open.result);
            open.onerror = () => reject(open.error);
          });
          const get = (store: string, key: string) =>
            new Promise<Uint8Array>((resolve, reject) => {
              const request = db.transaction(store).objectStore(store).get(key);
              request.onsuccess = () => resolve(request.result);
              request.onerror = () => reject(request.error);
            });
          const decode = (bytes: Uint8Array) =>
            JSON.parse(new TextDecoder().decode(bytes.subarray(1)));
          const head = decode(await get('heads', id))[1];
          const historyId = (direction === 'undo' ? head.past : head.future).at(-1) as string;
          const historyRaw = await get('records', historyId);
          const historyWrapper = decode(historyRaw);
          const history = historyWrapper[1];
          const metadataId = direction === 'undo' ? history.before : history.after;
          if (metadataId === head.meta) throw new Error('Fixture must target noncurrent metadata');
          const change = history.changes.find(
            (entry: any[]) => (direction === 'undo' ? entry[1] : entry[2])?.styles?.length,
          );
          if (!change) throw new Error('Fixture has no indexed historical style');
          const pointer = direction === 'undo' ? change[1] : change[2];
          const recordId = corruption === 'style' ? metadataId : historyId;
          const raw = corruption === 'style' ? await get('records', recordId) : historyRaw;
          const ledger = await get('sizes', recordId);
          const wrapper = corruption === 'style' ? decode(raw) : historyWrapper;
          if (corruption === 'style') {
            const index = wrapper[1].styleIds.indexOf(pointer.styles[0]);
            if (index < 0 || wrapper[1].styleIds.includes('missing_history_style'))
              throw new Error('Fixture style mismatch');
            wrapper[1].styleIds[index] = 'missing_history_style';
          } else change[0] = 'missing_sheet/0/0';
          const changed = new Uint8Array([0, ...new TextEncoder().encode(JSON.stringify(wrapper))]);
          const size = new Uint8Array([
            0,
            ...new TextEncoder().encode(JSON.stringify([recordId, changed.length])),
          ]);
          const tx = db.transaction(['records', 'sizes'], 'readwrite');
          const done = new Promise<void>((resolve, reject) => {
            tx.oncomplete = () => resolve();
            tx.onabort = () => reject(tx.error);
          });
          tx.objectStore('records').put(changed, recordId);
          tx.objectStore('sizes').put(size, recordId);
          await done;
          db.close();
          return { id, recordId, raw: Array.from(raw), ledger: Array.from(ledger) };
        },
        { direction, corruption },
      );
      await page.reload();
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      const capture = async () =>
        page.evaluate(async () => {
          const open = indexedDB.open('opensheet');
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            open.onsuccess = () => resolve(open.result);
            open.onerror = () => reject(open.error);
          });
          const names = ['heads', 'records', 'sizes', 'scratch'],
            tx = db.transaction(names);
          const result = await Promise.all(
            names.map(async (name) => {
              const store = tx.objectStore(name);
              const read = <T>(request: IDBRequest<T>) =>
                new Promise<T>((resolve, reject) => {
                  request.onsuccess = () => resolve(request.result);
                  request.onerror = () => reject(request.error);
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
      const attempt = await page.evaluate(async (direction) => {
        const book = window.opensheet.getWorkbook(),
          revision = book.revision;
        const error = await (direction === 'undo' ? book.undo() : book.redo()).then(
          () => ({ code: 'success', message: '' }),
          (reason: { code?: string; message?: string }) => ({
            code: reason.code,
            message: reason.message,
          }),
        );
        return {
          error,
          revision,
          after: book.revision,
          value: await book.getSheets()[0].range('A1').getValues(),
        };
      }, direction);
      expect(attempt).toMatchObject({
        error: { code: 'CORRUPT_STORAGE' },
        after: attempt.revision,
        value: [[372]],
      });
      expect(attempt.error.message).toContain('Block directory references an unknown');
      if (corruption === 'style')
        expect(attempt.error.message).toBe('Block directory references an unknown style');
      expect(await capture()).toEqual(before);
      await page.evaluate(async (fixture) => {
        const open = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          open.onsuccess = () => resolve(open.result);
          open.onerror = () => reject(open.error);
        });
        const tx = db.transaction(['records', 'sizes'], 'readwrite');
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve();
          tx.onabort = () => reject(tx.error);
        });
        tx.objectStore('records').put(new Uint8Array(fixture.raw), fixture.recordId);
        tx.objectStore('sizes').put(new Uint8Array(fixture.ledger), fixture.recordId);
        await done;
        db.close();
      }, fixture);
      const result = await page.evaluate(
        async ({ direction, corruption }) => {
          const app = window.opensheet;
          let book = app.getWorkbook();
          if (corruption === 'sheet') {
            await app.open(book.id);
            book = app.getWorkbook();
          }
          const applied = direction === 'undo' ? await book.undo() : await book.redo();
          const reversed = direction === 'undo' ? await book.redo() : await book.undo();
          await app.open(book.id);
          book = app.getWorkbook();
          const sheet = book.getSheets()[0];
          return {
            applied,
            reversed,
            value: await sheet.range('A1').getValues(),
            style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
          };
        },
        { direction, corruption },
      );
      expect(result.applied).toBe(true);
      expect(result.reversed).toBe(true);
      expect(result.value).toEqual([[372]]);
      expect(result.style.bold).toBe(true);
      expect(result.style.italic).toBe(direction === 'undo' ? true : undefined);
    });
  }
}
