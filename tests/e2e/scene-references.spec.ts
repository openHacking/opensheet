import { expect, test } from '@playwright/test';

test('reading scene references does not initialize or depend on legacy storage', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const path = '/src/scene-references.ts';
    const api = await import(/* @vite-ignore */ path);
    const key = `read-${Date.now()}`;
    const missing = await api.readSceneReference(key);
    const initialized = await api.initializeSceneReference(key, 'legacy-workbook');
    const original = Storage.prototype.getItem;
    Storage.prototype.getItem = () => {
      throw new DOMException('Denied', 'SecurityError');
    };
    try {
      const first = await api.readSceneReference(key);
      if (!first) throw new Error('Expected initialized reference');
      first.value = 'mutated';
      const second = await api.readSceneReference(key);
      return { missing, initialized, second };
    } finally {
      Storage.prototype.getItem = original;
    }
  });
  expect(result).toEqual({
    missing: null,
    initialized: { value: 'legacy-workbook', revision: 0 },
    second: { value: 'legacy-workbook', revision: 0 },
  });
});

test('transactional scene references serialize competing updates and reject ABA writes', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const other = await context.newPage();
  await other.goto('/#budget');
  await expect(other.locator('#scene-title')).toHaveText('Budget Simulator');
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

test('cancellation aborts a pending reference write but not an already committed write', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const path = '/src/scene-references.ts';
    const api = await import(/* @vite-ignore */ path);
    const key = `cancel-${Date.now()}`;
    const initial = await api.initializeSceneReference(key, 'old');
    const already = new AbortController();
    already.abort(new Error('cancel before open'));
    let early = '';
    try {
      await api.compareSceneReference(key, initial, 'ignored', { signal: already.signal });
    } catch (error) {
      early = (error as Error).message;
    }
    const newKey = `${key}-new`;
    let initialization = '';
    try {
      await api.initializeSceneReference(newKey, 'legacy', { signal: already.signal });
    } catch (error) {
      initialization = (error as Error).message;
    }
    const uninitialized = await api.readSceneReference(newKey);
    const controller = new AbortController();
    const put = IDBObjectStore.prototype.put;
    let reached = false;
    IDBObjectStore.prototype.put = function (...args) {
      const request = put.apply(this, args);
      if (this.name === 'references')
        request.addEventListener(
          'success',
          () => {
            reached = true;
            controller.abort(new Error('cancel pending write'));
          },
          { once: true },
        );
      return request;
    };
    let cancelled = '';
    try {
      await api.compareSceneReference(key, initial, 'cancelled', { signal: controller.signal });
    } catch (error) {
      cancelled = (error as Error).message;
    } finally {
      IDBObjectStore.prototype.put = put;
    }
    const afterAbort = await api.readSceneReference(key);
    const lateController = new AbortController();
    const get = IDBObjectStore.prototype.get;
    let lateReached = false;
    IDBObjectStore.prototype.get = function (...args) {
      if (this.name === 'references')
        this.transaction.addEventListener(
          'complete',
          () => {
            lateReached = true;
            lateController.abort(new Error('late abort'));
          },
          { once: true },
        );
      return get.apply(this, args);
    };
    let committed;
    try {
      committed = await api.compareSceneReference(key, initial, 'saved', {
        signal: lateController.signal,
      });
    } finally {
      IDBObjectStore.prototype.get = get;
    }
    const afterCommit = await api.readSceneReference(key);
    return {
      early,
      initialization,
      uninitialized,
      reached,
      cancelled,
      afterAbort,
      lateReached,
      committed,
      afterCommit,
    };
  });
  expect(result).toEqual({
    early: 'cancel before open',
    initialization: 'cancel before open',
    uninitialized: null,
    reached: true,
    cancelled: 'cancel pending write',
    afterAbort: { value: 'old', revision: 0 },
    lateReached: true,
    committed: { updated: true, current: { value: 'saved', revision: 1 } },
    afterCommit: { value: 'saved', revision: 1 },
  });
});

test('cancelling a queued reference transaction does not publish after the blocker releases', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const path = '/src/scene-references.ts';
    const api = await import(/* @vite-ignore */ path);
    const key = `queued-${Date.now()}`;
    const initial = await api.initializeSceneReference(key, 'old');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const open = indexedDB.open('opensheet-playground-references', 1);
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('references', 'readwrite');
    const store = tx.objectStore('references');
    let release = false;
    let held!: () => void;
    const started = new Promise<void>((resolve) => (held = resolve));
    const keepAlive = () => {
      const request = store.get(key);
      request.onsuccess = () => {
        held();
        if (!release) keepAlive();
      };
    };
    const completed = new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    keepAlive();
    await started;
    const controller = new AbortController();
    const transaction = IDBDatabase.prototype.transaction;
    let queued!: () => void;
    const queuedPromise = new Promise<void>((resolve) => (queued = resolve));
    IDBDatabase.prototype.transaction = function (...args) {
      const next = transaction.apply(this, args);
      if (this.name === 'opensheet-playground-references' && args[1] === 'readwrite') queued();
      return next;
    };
    const pending = api.compareSceneReference(key, initial, 'new', { signal: controller.signal });
    await queuedPromise;
    IDBDatabase.prototype.transaction = transaction;
    controller.abort(new Error('cancel queued write'));
    release = true;
    const outcome = await pending.then(
      () => 'unexpected success',
      (error: Error) => error.message,
    );
    await completed;
    db.close();
    return { outcome, reference: await api.readSceneReference(key) };
  });
  expect(result).toEqual({
    outcome: 'cancel queued write',
    reference: { value: 'old', revision: 0 },
  });
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
    let readCorruption = '';
    try {
      await api.readSceneReference(key);
    } catch (error) {
      readCorruption = (error as Error).message;
    }
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
    return { corruption, readCorruption, afterCorrupt, limit, afterExhausted };
  });
  expect(result).toEqual({
    corruption: 'Corrupt scene reference',
    readCorruption: 'Corrupt scene reference',
    afterCorrupt: { value: 'old', revision: 4, extra: true },
    limit: 'Scene reference revision exhausted',
    afterExhausted: { value: 'old', revision: Number.MAX_SAFE_INTEGER },
  });
});
