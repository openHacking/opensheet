import { expect, test } from '@playwright/test';

for (const variant of [
  'missing-page',
  'missing-ledger',
  'invalid-ledger',
  'wrong-size',
  'truncated-page',
] as const) {
  test(`${variant} stops recovery before any record or scratch cleanup`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const fixture = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[987]]);
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
        JSON.parse(new TextDecoder().decode(bytes.subarray(1)))[1];
      const encode = (key: string, value: unknown) => {
        const bytes = new TextEncoder().encode(JSON.stringify([key, value]));
        const result = new Uint8Array(bytes.length + 1);
        result.set(bytes, 1);
        return result;
      };
      const head = decode(await get('heads', id));
      const directoryId = Object.values(head.directory)[0] as string;
      const directory = decode(await get('records', directoryId));
      const pageId = (Object.values(directory)[0] as { id: string }).id;
      const orphan = `${id}/zz-recovery-orphan`;
      const scratch = `${id}/zz-recovery-scratch`;
      const tx = db.transaction(['records', 'sizes', 'scratch'], 'readwrite');
      tx.objectStore('records').put(new Uint8Array([0, 42]), orphan);
      tx.objectStore('sizes').put(encode(orphan, kind === 'invalid-ledger' ? -1 : 2), orphan);
      tx.objectStore('scratch').put(new Uint8Array([0, 43]), scratch);
      if (kind === 'missing-page') tx.objectStore('records').delete(pageId);
      if (kind === 'missing-ledger') tx.objectStore('sizes').delete(pageId);
      if (kind === 'wrong-size') tx.objectStore('sizes').put(encode(pageId, 1), pageId);
      if (kind === 'truncated-page') tx.objectStore('records').put(new Uint8Array([1]), pageId);
      await new Promise<void>((resolve, reject) => {
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      const snapshot = async () => {
        const result: Record<string, unknown> = {};
        for (const store of ['heads', 'records', 'sizes', 'scratch']) {
          const transaction = db.transaction(store);
          const keyRequest = transaction.objectStore(store).getAllKeys();
          const valueRequest = transaction.objectStore(store).getAll();
          const [keys, values] = await Promise.all(
            [keyRequest, valueRequest].map(
              (request) =>
                new Promise<any[]>((resolve, reject) => {
                  request.onsuccess = () => resolve(request.result);
                  request.onerror = () => reject(request.error);
                }),
            ),
          );
          result[store] = keys.map((key, index) => [
            String(key),
            Array.from(values[index] as Uint8Array),
          ]);
        }
        return result;
      };
      const before = await snapshot();
      db.close();
      return { id, before };
    }, variant);
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await expect(recovery.locator('#startup-recovery-message')).toContainText('CORRUPT_STORAGE');
    expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
      fixture.id,
    );
    const after = await recovery.evaluate(async () => {
      const open = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
      });
      const result: Record<string, unknown> = {};
      for (const store of ['heads', 'records', 'sizes', 'scratch']) {
        const transaction = db.transaction(store);
        const requests = [
          transaction.objectStore(store).getAllKeys(),
          transaction.objectStore(store).getAll(),
        ];
        const [keys, values] = await Promise.all(
          requests.map(
            (request) =>
              new Promise<any[]>((resolve, reject) => {
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error);
              }),
          ),
        );
        result[store] = keys.map((key, index) => [
          String(key),
          Array.from(values[index] as Uint8Array),
        ]);
      }
      db.close();
      return result;
    });
    expect(after).toEqual(fixture.before);
    await recovery.close();
  });
}
