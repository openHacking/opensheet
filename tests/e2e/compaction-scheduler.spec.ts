import { expect, test } from '@playwright/test';

test('idle compaction eventually runs and preserves committed data across reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const supported = await page.evaluate(
    () => typeof CompressionStream !== 'undefined' && typeof DecompressionStream !== 'undefined',
  );
  test.skip(!supported, 'Browser does not expose compression streams');
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const values = Array.from({ length: 64 }, () =>
      Array.from({ length: 12 }, () => 'repeated payload for compression'),
    );
    await sheet.range('A1:L64').setValues(values);
    return { id: book.id, revision: book.revision };
  });
  await expect
    .poll(
      () =>
        page.evaluate(async (id) => {
          const open = indexedDB.open('opensheet');
          const db = await new Promise<IDBDatabase>((resolve, reject) => {
            open.onsuccess = () => resolve(open.result);
            open.onerror = () => reject(open.error);
          });
          const request = db
            .transaction('records')
            .objectStore('records')
            .getAll(IDBKeyRange.bound(`${id}/`, `${id}/\uffff`));
          const values = await new Promise<Uint8Array[]>((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
          });
          db.close();
          return values.some(
            (value) =>
              value[0] === 1 &&
              value[1] === 79 &&
              value[2] === 83 &&
              value[3] === 71 &&
              value[4] === 3,
          );
        }, before.id),
      { timeout: 15000 },
    )
    .toBe(true);
  await page.evaluate(async (id) => window.opensheet.open(id), before.id);
  const after = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const values = await book.getSheets()[0].range('A1:L64').getValues();
    return {
      revision: book.revision,
      count: values.flat().filter((value) => value === 'repeated payload for compression').length,
      undo: await book.undo(),
    };
  });
  expect(after).toEqual({ revision: before.revision, count: 768, undo: true });
});
