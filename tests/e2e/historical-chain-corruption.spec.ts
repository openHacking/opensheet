import { expect, test } from '@playwright/test';

for (const direction of ['undo', 'redo'] as const) {
  for (const corruption of ['page', 'metadata', 'extent', 'style-index', 'row-mask'] as const) {
    for (const depth of ['next', 'later'] as const) {
      test(`rejects a disconnected ${direction} ${corruption} history at ${depth} depth before cleanup`, async ({
        page,
        context,
      }) => {
        await page.goto('/#budget');
        await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
        const fixture = await page.evaluate(
          async ({ direction, corruption, depth }) => {
            const book = window.opensheet.getWorkbook(),
              sheet = book.getSheets()[0];
            await sheet.range('A1').setValues([[372]]);
            if (corruption !== 'metadata') {
              await sheet.range('A1').setValues([[373]]);
              await sheet.range('A1').setValues([[374]]);
            } else {
              await sheet.range('A1').setStyle({ bold: true });
              await sheet.range('A1').setStyle({ italic: true });
            }
            if (depth === 'later') {
              if (corruption !== 'metadata') await sheet.range('A2').setValues([[41]]);
              else await sheet.range('A2').setStyle({ underline: true });
            }
            if (direction === 'redo') {
              await book.undo();
              if (depth === 'later') await book.undo();
            }
            const id = book.id;
            await book.close();
            const request = indexedDB.open('opensheet');
            const db = await new Promise<IDBDatabase>((resolve, reject) => {
              request.onsuccess = () => resolve(request.result);
              request.onerror = () => reject(request.error);
            });
            const read = (store: string, key: string) =>
              new Promise<Uint8Array>((resolve, reject) => {
                const req = db.transaction(store).objectStore(store).get(key);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
              });
            const decode = (raw: Uint8Array) =>
              JSON.parse(new TextDecoder().decode(raw.subarray(1)));
            const head = decode(await read('heads', id))[1];
            const historyId = (direction === 'undo' ? head.past : head.future).at(
              depth === 'later' ? -2 : -1,
            );
            const raw = await read('records', historyId),
              ledger = await read('sizes', historyId);
            const wrapper = decode(raw),
              history = wrapper[1];
            if (corruption !== 'metadata') {
              const change = history.changes.find(
                (entry: any[]) => entry[1] && entry[2] && entry[1].id !== entry[2].id,
              );
              if (!change) throw new Error('Fixture needs two distinct page versions');
              if (corruption === 'page') {
                if (direction === 'undo') change[2] = change[1];
                else change[1] = change[2];
              } else {
                const source = direction === 'undo' ? change[2] : change[1];
                if (corruption === 'extent') source.endRow = source.endRow === 1 ? 2 : 1;
                if (corruption === 'row-mask') source.rowMask = source.rowMask === '1' ? '2' : '1';
                if (corruption === 'style-index') {
                  if (!source.styles?.length) throw new Error('Fixture needs indexed styles');
                  source.styles = source.styles.slice(1);
                }
              }
            } else {
              if (history.before === history.after)
                throw new Error('Fixture needs two metadata versions');
              if (direction === 'undo') history.after = history.before;
              else history.before = history.after;
            }
            const altered = new Uint8Array([
              0,
              ...new TextEncoder().encode(JSON.stringify(wrapper)),
            ]);
            const size = new Uint8Array([
              0,
              ...new TextEncoder().encode(JSON.stringify([historyId, altered.length])),
            ]);
            const tx = db.transaction(['records', 'sizes', 'scratch'], 'readwrite');
            const done = new Promise<void>((resolve, reject) => {
              tx.oncomplete = () => resolve();
              tx.onabort = () => reject(tx.error);
            });
            tx.objectStore('records').put(altered, historyId);
            tx.objectStore('sizes').put(size, historyId);
            tx.objectStore('scratch').put(new Uint8Array([0, 42]), `${id}/history-chain-sentinel`);
            await done;
            db.close();
            return { id, historyId, raw: Array.from(raw), ledger: Array.from(ledger) };
          },
          { direction, corruption, depth },
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
                const read = <T>(request: IDBRequest<T>) =>
                  new Promise<T>((resolve, reject) => {
                    request.onsuccess = () => resolve(request.result);
                    request.onerror = () => reject(request.error);
                  });
                const [keys, values] = await Promise.all([
                  read(store.getAllKeys()),
                  read(store.getAll()),
                ]);
                return [name, keys.map((key, i) => [key, Array.from(values[i] as Uint8Array)])];
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
          'Workbook history does not match the current revision',
        );
        expect(await capture()).toEqual(before);
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
          tx.objectStore('records').put(new Uint8Array(fixture.raw), fixture.historyId);
          tx.objectStore('sizes').put(new Uint8Array(fixture.ledger), fixture.historyId);
          await done;
          db.close();
        }, fixture);
        await recovery.getByRole('button', { name: 'Try again' }).click();
        await expect(
          recovery.getByRole('heading', { name: 'Unable to open workbook' }),
        ).toBeHidden();
        const result = await recovery.evaluate(async (direction) => {
          const app = window.opensheet;
          let book = app.getWorkbook();
          const first = direction === 'undo' ? await book.undo() : await book.redo();
          const second = direction === 'undo' ? await book.redo() : await book.undo();
          await app.open(book.id);
          book = app.getWorkbook();
          return { first, second, value: await book.getSheets()[0].range('A1').getValues() };
        }, direction);
        expect(result).toEqual({
          first: true,
          second: true,
          value: [[corruption !== 'metadata' ? (direction === 'undo' ? 374 : 373) : 372]],
        });
      });
    }
  }
}
