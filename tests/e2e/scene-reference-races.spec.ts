import { expect, test } from '@playwright/test';

for (const operation of ['initial', 'reset'] as const) {
  test(`${operation} does not overwrite a scene reference changed during preparation`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const original = await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const prototype = Object.getPrototypeOf(book);
      const ready = prototype.ready;
      prototype.ready = async function () {
        const result = await ready.call(this);
        if (this !== book) {
          prototype.ready = ready;
          (window as any).preparedSceneId = this.id;
          await new Promise<void>((resolve) => ((window as any).resumeSceneReference = resolve));
        }
        return result;
      };
      return { id: book.id, revision: book.revision };
    });
    if (operation === 'initial') await page.locator('#scene-select').selectOption('sales');
    else await page.getByRole('button', { name: 'Reset demo' }).click();
    await expect.poll(() => page.evaluate(() => !!(window as any).resumeSceneReference)).toBe(true);
    const candidate = await page.evaluate(() => (window as any).preparedSceneId as string);
    const other = await context.newPage();
    await other.goto('/#budget');
    await expect(other.locator('#scene-select')).toBeEnabled();
    const key = `opensheet.scene.${operation === 'initial' ? 'sales' : 'budget'}`;
    const otherId = await other.evaluate(async (key) => {
      const book = window.opensheet.getWorkbook();
      const snapshot = await book.toJSON();
      const Workbook = book.constructor as new (file: typeof snapshot) => typeof book;
      const replacement = new Workbook({ ...snapshot, workbookId: `other_${crypto.randomUUID()}` });
      await replacement.ready();
      await replacement
        .getSheets()[0]
        .range('A1')
        .setValues([[8642]]);
      await replacement.close();
      localStorage.setItem(key, replacement.id);
      return replacement.id;
    }, key);
    await expect.poll(() => page.evaluate((key) => localStorage.getItem(key), key)).toBe(otherId);
    await page.evaluate(() => (window as any).resumeSceneReference());
    await expect(page.locator('#notice')).toContainText('Saved scene changed');
    const result = await page.evaluate(
      async ({ key, candidate }) => {
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const names = ['heads', 'records', 'sizes', 'scratch'];
        const tx = db.transaction(names);
        const candidateCounts = await Promise.all(
          names.map(
            (name) =>
              new Promise<number>((resolve, reject) => {
                const range =
                  name === 'heads'
                    ? IDBKeyRange.only(candidate)
                    : IDBKeyRange.bound(`${candidate}/`, `${candidate}/\uffff`);
                const read = tx.objectStore(name).count(range);
                read.onsuccess = () => resolve(read.result);
                read.onerror = () => reject(read.error);
              }),
          ),
        );
        db.close();
        const book = window.opensheet.getWorkbook();
        return {
          id: book.id,
          revision: book.revision,
          reference: localStorage.getItem(key),
          candidateCounts,
        };
      },
      { key, candidate },
    );
    expect(result).toEqual({ ...original, reference: otherId, candidateCounts: [0, 0, 0, 0] });
    if (operation === 'initial') {
      await page.locator('#scene-select').selectOption('sales');
      await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
    } else {
      await page.reload();
      await expect(page.locator('#scene-select')).toBeEnabled();
    }
    expect(
      await page.evaluate(async () => ({
        id: window.opensheet.getWorkbook().id,
        value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      })),
    ).toEqual({ id: otherId, value: [[8642]] });
    await other.close();
  });
}

for (const operation of ['initial', 'reset'] as const) {
  test(`${operation} preserves the current workbook if the final reference read fails`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const original = await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const prototype = Object.getPrototypeOf(book);
      const ready = prototype.ready;
      prototype.ready = async function () {
        const result = await ready.call(this);
        if (this !== book) {
          prototype.ready = ready;
          (window as any).preparedSceneId = this.id;
          await new Promise<void>((resolve) => ((window as any).resumeSceneReference = resolve));
        }
        return result;
      };
      return { id: book.id, revision: book.revision };
    });
    if (operation === 'initial') await page.locator('#scene-select').selectOption('sales');
    else await page.getByRole('button', { name: 'Reset demo' }).click();
    await expect.poll(() => page.evaluate(() => !!(window as any).resumeSceneReference)).toBe(true);
    const key = `opensheet.scene.${operation === 'initial' ? 'sales' : 'budget'}`;
    const reference = await page.evaluate((key) => localStorage.getItem(key), key);
    await page.evaluate((key) => {
      const get = Storage.prototype.getItem;
      Storage.prototype.getItem = function (name) {
        if (this === localStorage && name === key) {
          Storage.prototype.getItem = get;
          throw new DOMException('Reference read denied', 'SecurityError');
        }
        return get.call(this, name);
      };
      (window as any).resumeSceneReference();
    }, key);
    await expect(page.locator('#notice')).toContainText('Reference read denied');
    const result = await page.evaluate(async (key) => {
      const candidate = (window as any).preparedSceneId as string;
      const request = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const read = db.transaction('heads').objectStore('heads').get(candidate);
      const head = await new Promise<unknown>((resolve, reject) => {
        read.onsuccess = () => resolve(read.result);
        read.onerror = () => reject(read.error);
      });
      db.close();
      const book = window.opensheet.getWorkbook();
      return {
        id: book.id,
        revision: book.revision,
        reference: localStorage.getItem(key),
        candidateExists: head !== undefined,
      };
    }, key);
    expect(result).toEqual({ ...original, reference, candidateExists: false });
    if (operation === 'initial') {
      await page.locator('#scene-select').selectOption('sales');
      await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
    } else {
      await page.getByRole('button', { name: 'Reset demo' }).click();
      await expect(page.getByRole('button', { name: 'Reset demo' })).toBeEnabled();
      expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).not.toBe(original.id);
    }
  });
}
