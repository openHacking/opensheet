import { expect, test } from '@playwright/test';

for (const corruption of ['invalid-utf8', 'wrong-key', 'null-head'] as const) {
  test(`corrupt ${corruption} storage head preserves records and supports repair`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const id = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[823]]);
      return book.id;
    });
    const capture = async () =>
      page.evaluate(async () => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const names = ['heads', 'records', 'sizes', 'scratch'];
        const tx = db.transaction(names, 'readonly');
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
            return [name, keys.map((key, index) => [key, Array.from(values[index] as Uint8Array)])];
          }),
        );
        db.close();
        return result;
      });
    const before = await capture();
    const original = await page.evaluate(
      async ({ id, corruption }) => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const tx = db.transaction('heads', 'readwrite'),
          store = tx.objectStore('heads');
        const prior = await new Promise<Uint8Array>((resolve, reject) => {
          const read = store.get(id);
          read.onsuccess = () => resolve(read.result);
          read.onerror = () => reject(read.error);
        });
        const bytes =
          corruption === 'invalid-utf8'
            ? new Uint8Array([0, 0x5b, 0x22, 0xff, 0x22, 0x2c, 0x31, 0x5d])
            : new Uint8Array([
                0,
                ...new TextEncoder().encode(
                  JSON.stringify(
                    corruption === 'null-head' ? [id, null] : ['other-id', { revision: 0 }],
                  ),
                ),
              ]);
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve();
          tx.onabort = () => reject(tx.error);
        });
        store.put(bytes, id);
        await done;
        db.close();
        return Array.from(prior);
      },
      { id, corruption },
    );
    const corrupted = await capture();
    await page.reload();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await expect(page.locator('#startup-recovery-message')).toContainText('CORRUPT_STORAGE');
    expect(await capture()).toEqual(corrupted);
    expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
    await page.evaluate(
      async ({ id, original }) => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const tx = db.transaction('heads', 'readwrite');
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve();
          tx.onabort = () => reject(tx.error);
        });
        tx.objectStore('heads').put(new Uint8Array(original), id);
        await done;
        db.close();
      },
      { id, original },
    );
    expect(await capture()).toEqual(before);
    await page.getByRole('button', { name: 'Try again' }).click();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[823]]);
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book.undo();
      await book.redo();
      await window.opensheet.open(book.id);
    });
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[823]]);
  });
}
