import { expect, test } from '@playwright/test';

test('a visible reader catches up after missing edit and deletion notifications', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const writer = window.opensheet.getWorkbook();
    const Constructor = writer.constructor as new (
      snapshot: undefined,
      options: object,
      id: string,
    ) => typeof writer;
    const reader = await new Constructor(undefined, {}, writer.id).ready();
    ((reader as any).channel as BroadcastChannel).onmessage = null;
    const sheet = reader.getSheets()[0];
    await reader.prefetch(sheet.id, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    await writer
      .getSheets()[0]
      .range('A1')
      .setValues([[12345]]);
    const stale = reader.revision < writer.revision;
    window.dispatchEvent(new Event('focus'));
    const until = async (condition: () => boolean) => {
      const deadline = Date.now() + 5000;
      while (!condition() && Date.now() < deadline)
        await new Promise((resolve) => setTimeout(resolve, 20));
      return condition();
    };
    const caughtUp = await until(() => reader.revision === writer.revision);
    const value = (await sheet.range('A1').getValues())[0][0];
    await reader.prefetch(sheet.id, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    await writer.deleteStorage();
    document.dispatchEvent(new Event('visibilitychange'));
    const failed = await until(() => reader.persistenceState.state === 'failed');
    const error = reader.persistenceState.error?.code;
    const cached = reader.peekValue(sheet.id, 0, 0);
    const worker = (reader as any).worker as Worker;
    const original = worker.postMessage.bind(worker);
    let requests = 0;
    worker.postMessage = ((message: any) => {
      requests++;
      original(message);
    }) as typeof worker.postMessage;
    reader.dispose();
    window.dispatchEvent(new Event('focus'));
    document.dispatchEvent(new Event('visibilitychange'));
    await new Promise((resolve) => setTimeout(resolve, 50));
    return { stale, caughtUp, value, failed, error, cached, requests };
  });
  expect(result).toEqual({
    stale: true,
    caughtUp: true,
    value: 12345,
    failed: true,
    error: 'CORRUPT_STORAGE',
    cached: { error: '#LOADING' },
    requests: 0,
  });
});

test('read-only opens remain usable while another worker publishes edits', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const Constructor = book.constructor as new (
      snapshot: undefined,
      options: object,
      id: string,
    ) => typeof book;
    const follower = await new Constructor(undefined, {}, book.id).ready();
    const writes = (async () => {
      for (let i = 0; i < 40; i++)
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([[i]]);
    })();
    const failures: string[] = [];
    const values: unknown[] = [];
    for (let i = 0; i < 20; i++) {
      let reader: typeof book | undefined;
      try {
        reader = await new Constructor(undefined, {}, book.id).ready();
        values.push((await reader.getSheets()[0].range('A1').getValues())[0][0]);
      } catch (error) {
        failures.push(`${(error as { code?: string }).code ?? ''}: ${(error as Error).message}`);
      } finally {
        await reader?.close();
      }
    }
    await writes;
    const deadline = Date.now() + 5000;
    while (follower.revision !== book.revision && Date.now() < deadline)
      await new Promise((resolve) => setTimeout(resolve, 20));
    const followerRevision = follower.revision;
    const writerRevision = book.revision;
    const followerValue = (await follower.getSheets()[0].range('A1').getValues())[0][0];
    await follower.close();
    return {
      failures,
      values,
      final: (await book.getSheets()[0].range('A1').getValues())[0][0],
      followerRevision,
      writerRevision,
      followerValue,
    };
  });
  expect(result.failures).toEqual([]);
  expect(result.values).toHaveLength(20);
  expect(result.final).toBe(39);
  expect(result.followerRevision).toBe(result.writerRevision);
  expect(result.followerValue).toBe(39);
});

test('a stable corrupt head still fails explicitly for a read-only opener', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const Constructor = book.constructor as new (
      snapshot: undefined,
      options: object,
      id: string,
    ) => typeof book;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[77]]);
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
    const head = decode(await get('heads', book.id));
    const directory = decode(await get('records', Object.values(head.directory)[0] as string));
    const pageId = (Object.values(directory)[0] as { id: string }).id;
    const tx = db.transaction('sizes', 'readwrite');
    tx.objectStore('sizes').delete(pageId);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    try {
      const reader = await new Constructor(undefined, {}, book.id).ready();
      await reader.close();
      return 'unexpected success';
    } catch (error) {
      return (error as { code?: string }).code ?? '';
    }
  });
  expect(result).toBe('CORRUPT_STORAGE');
});

test('coalesces notification bursts and still converges to the latest revision', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const Constructor = book.constructor as new (
      snapshot: undefined,
      options: object,
      id: string,
    ) => typeof book;
    const reader = await new Constructor(undefined, {}, book.id).ready();
    const worker = (reader as any).worker as Worker;
    const original = worker.postMessage.bind(worker);
    let requests = 0;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'reloadIfChanged') requests++;
      return original(message, ...rest);
    }) as typeof worker.postMessage;
    const channel = new BroadcastChannel(`opensheet:opensheet:${book.id}`);
    try {
      for (let i = 0; i < 100; i++) channel.postMessage({ synthetic: i });
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[2468]]);
      const deadline = Date.now() + 5000;
      while (reader.revision !== book.revision && Date.now() < deadline)
        await new Promise((resolve) => setTimeout(resolve, 20));
      await new Promise((resolve) => setTimeout(resolve, 250));
      return {
        requests,
        readerRevision: reader.revision,
        writerRevision: book.revision,
        value: (await reader.getSheets()[0].range('A1').getValues())[0][0],
      };
    } finally {
      channel.close();
      await reader.close();
    }
  });
  expect(result.readerRevision).toBe(result.writerRevision);
  expect(result.value).toBe(2468);
  expect(result.requests).toBeLessThan(20);
});
