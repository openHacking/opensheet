import { expect, test } from '@playwright/test';

for (const corruption of ['collision', 'oversized', 'missing'] as const) {
  test(`invalid stored style ${corruption} preserves the workbook and can be repaired`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const id = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      await sheet.range('A1:A2').setValues([[17], [29]]);
      await sheet.range('A1').setStyle({ bold: true });
      await sheet.range('A2').setStyle({ italic: true });
      return book.id;
    });
    const capture = async () =>
      page.evaluate(async () => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const names = ['heads', 'records', 'sizes', 'scratch'],
          tx = db.transaction(names, 'readonly');
        const entries = await Promise.all(
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
        return entries;
      });
    const before = await capture();
    const original = await page.evaluate(
      async ({ id, corruption }) => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const tx = db.transaction(['heads', 'records', 'sizes'], 'readwrite');
        const read = <T>(req: IDBRequest<T>) =>
          new Promise<T>((resolve, reject) => {
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
          });
        const head = (await read(tx.objectStore('heads').get(id))) as Uint8Array;
        const storedHead = JSON.parse(new TextDecoder().decode(head.subarray(1)))[1];
        const metaId = storedHead.meta as string;
        const [raw, ledger] = (await Promise.all([
          read(tx.objectStore('records').get(metaId)),
          read(tx.objectStore('sizes').get(metaId)),
        ])) as Uint8Array[];
        const wrapper = JSON.parse(new TextDecoder().decode(raw.subarray(1)));
        if (wrapper[1].styleIds.length < 2) throw new Error('Expected at least two styles');
        if (corruption === 'collision') {
          wrapper[1].styleIds[0] = 0;
          wrapper[1].styleIds[1] = '0';
        } else if (corruption === 'oversized') wrapper[1].styleIds[0] = 'a'.repeat(101);
        else {
          const directories = (await Promise.all(
            (Object.values(storedHead.directory) as string[]).map((id) =>
              read(tx.objectStore('records').get(id)),
            ),
          )) as Uint8Array[];
          const referenced = directories
            .flatMap(
              (bytes) =>
                Object.values(JSON.parse(new TextDecoder().decode(bytes.subarray(1)))[1]) as {
                  styles?: string[];
                }[],
            )
            .flatMap((pointer) => pointer.styles ?? [])
            .find((id) => wrapper[1].styleIds.includes(id));
          if (!referenced) throw new Error('Fixture has no indexed style');
          wrapper[1].styleIds[wrapper[1].styleIds.indexOf(referenced)] = 'missing_style';
        }
        const corrupt = new Uint8Array([0, ...new TextEncoder().encode(JSON.stringify(wrapper))]);
        const size = new Uint8Array([
          0,
          ...new TextEncoder().encode(JSON.stringify([metaId, corrupt.length])),
        ]);
        const done = new Promise<void>((resolve, reject) => {
          tx.oncomplete = () => resolve();
          tx.onabort = () => reject(tx.error);
        });
        tx.objectStore('records').put(corrupt, metaId);
        tx.objectStore('sizes').put(size, metaId);
        await done;
        db.close();
        return { metaId, raw: Array.from(raw), ledger: Array.from(ledger) };
      },
      { id, corruption },
    );
    const corrupt = await capture();
    await page.reload();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await expect(page.locator('#startup-recovery-message')).toContainText(
      corruption === 'missing'
        ? 'Block directory references an unknown style'
        : 'Invalid stored style identities',
    );
    expect(await capture()).toEqual(corrupt);
    expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
    await page.evaluate(async (original) => {
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
      tx.objectStore('records').put(new Uint8Array(original.raw), original.metaId);
      tx.objectStore('sizes').put(new Uint8Array(original.ledger), original.metaId);
      await done;
      db.close();
    }, original);
    expect(await capture()).toEqual(before);
    await page.getByRole('button', { name: 'Try again' }).click();
    await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
    const result = await page.evaluate(async () => {
      const app = window.opensheet;
      let book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      await book.undo();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      return {
        values: await sheet.range('A1:A2').getValues(),
        styles: [
          book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
          book.getStyle((await book.getCell(sheet.id, 1, 0))?.styleId),
        ],
      };
    });
    expect(result.values).toEqual([[17], [29]]);
    expect(result.styles[0].bold).toBe(true);
    expect(result.styles[1].italic).toBe(true);
  });
}
