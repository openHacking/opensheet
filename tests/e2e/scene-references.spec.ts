import { expect, test } from '@playwright/test';

test('transactional scene references serialize competing updates and reject ABA writes', async ({
  page,
  context,
}) => {
  const other = await context.newPage();
  await Promise.all([page.goto('/#budget'), other.goto('/#budget')]);
  await Promise.all([
    expect(page.locator('#scene-title')).toHaveText('Budget Simulator'),
    expect(other.locator('#scene-title')).toHaveText('Budget Simulator'),
  ]);
  const key = `scene-${Date.now()}`;
  const initialize = (p: typeof page, legacy: string | null) =>
    p.evaluate(
      async ({ key, legacy }) => {
        const path = '/src/scene-references.ts';
        const api = await import(/* @vite-ignore */ path);
        return api.initializeSceneReference(key, legacy);
      },
      { key, legacy },
    );
  const [first, second] = await Promise.all([
    initialize(page, 'old-a'),
    initialize(other, 'old-b'),
  ]);
  expect(first).toEqual(second);
  expect(['old-a', 'old-b']).toContain(first.value);
  expect(first.revision).toBe(0);
  const swap = (p: typeof page, expected: typeof first, next: string | null) =>
    p.evaluate(
      async ({ key, expected, next }) => {
        const path = '/src/scene-references.ts';
        const api = await import(/* @vite-ignore */ path);
        return api.compareSceneReference(key, expected, next);
      },
      { key, expected, next },
    );
  const [a, b] = await Promise.all([swap(page, first, 'new-a'), swap(other, first, 'new-b')]);
  expect([a.updated, b.updated].filter(Boolean)).toHaveLength(1);
  const current = await initialize(page, null);
  expect(current.revision).toBe(1);
  expect(['new-a', 'new-b']).toContain(current.value);
  const back = await swap(other, current, first.value);
  expect(back).toEqual({ updated: true, current: { value: first.value, revision: 2 } });
  expect(await swap(page, first, 'stale')).toEqual({ updated: false, current: back.current });
  await other.reload();
  expect(await initialize(other, 'ignored')).toEqual(back.current);
  await other.close();
});

test('aborted scene reference writes preserve the previous record', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const path = '/src/scene-references.ts';
    const api = await import(/* @vite-ignore */ path);
    const key = `abort-${Date.now()}`;
    const initial = await api.initializeSceneReference(key, 'old');
    const put = IDBObjectStore.prototype.put;
    let reached = false;
    IDBObjectStore.prototype.put = function (...args) {
      const request = put.apply(this, args);
      if (this.name === 'references')
        request.addEventListener(
          'success',
          () => {
            reached = true;
            this.transaction.abort();
          },
          { once: true },
        );
      return request;
    };
    let rejected = false;
    try {
      await api.compareSceneReference(key, initial, 'new');
    } catch {
      rejected = true;
    } finally {
      IDBObjectStore.prototype.put = put;
    }
    return { reached, rejected, after: await api.initializeSceneReference(key, null) };
  });
  expect(result).toEqual({ reached: true, rejected: true, after: { value: 'old', revision: 0 } });
});

test('corrupt or exhausted scene reference records are rejected without rewriting them', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const path = '/src/scene-references.ts';
    const api = await import(/* @vite-ignore */ path);
    const key = `broken-${Date.now()}`;
    await api.initializeSceneReference(key, null);
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('opensheet-playground-references', 1);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const write = (value: unknown) =>
      new Promise<void>((resolve, reject) => {
        const tx = db.transaction('references', 'readwrite');
        tx.objectStore('references').put(value, key);
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
    const read = () =>
      new Promise<unknown>((resolve, reject) => {
        const request = db.transaction('references').objectStore('references').get(key);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const corrupt = { value: 'old', revision: 4, extra: true };
    await write(corrupt);
    let corruption = '';
    try {
      await api.initializeSceneReference(key, 'fallback');
    } catch (error) {
      corruption = (error as Error).message;
    }
    const afterCorrupt = await read();
    const exhausted = { value: 'old', revision: Number.MAX_SAFE_INTEGER };
    await write(exhausted);
    let limit = '';
    try {
      await api.compareSceneReference(key, exhausted, 'new');
    } catch (error) {
      limit = (error as Error).message;
    }
    const afterExhausted = await read();
    db.close();
    return { corruption, afterCorrupt, limit, afterExhausted };
  });
  expect(result).toEqual({
    corruption: 'Corrupt scene reference',
    afterCorrupt: { value: 'old', revision: 4, extra: true },
    limit: 'Scene reference revision exhausted',
    afterExhausted: { value: 'old', revision: Number.MAX_SAFE_INTEGER },
  });
});
