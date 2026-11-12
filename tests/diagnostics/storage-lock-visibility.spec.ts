import { expect, test } from '@playwright/test';

test('a waiting Web Lock holder observes the previous holder localStorage write', async ({
  page,
  context,
}) => {
  const other = await context.newPage();
  await Promise.all([page.goto('/#budget'), other.goto('/#budget')]);
  await Promise.all([
    expect(page.locator('#scene-title')).toHaveText('Budget Simulator'),
    expect(other.locator('#scene-title')).toHaveText('Budget Simulator'),
  ]);
  const name = `opensheet.test.visibility.${Date.now()}-${Math.random()}`;
  await page.evaluate((name) => {
    let release!: () => void;
    const pending = new Promise<void>((resolve) => {
      release = resolve;
    });
    (window as any).releaseVisibilityLock = release;
    void navigator.locks.request(name, async () => {
      (window as any).visibilityLockHeld = true;
      await pending;
      localStorage.setItem(name, 'published');
    });
  }, name);
  await expect.poll(() => page.evaluate(() => (window as any).visibilityLockHeld)).toBe(true);
  await other.evaluate((name) => {
    void navigator.locks.request(name, () => {
      (window as any).visibilityObserved = localStorage.getItem(name);
      (window as any).visibilityDone = true;
    });
  }, name);
  await expect
    .poll(() =>
      page.evaluate(
        async (name) => (await navigator.locks.query()).pending?.some((lock) => lock.name === name),
        name,
      ),
    )
    .toBe(true);
  await page.evaluate(() => (window as any).releaseVisibilityLock());
  await expect.poll(() => other.evaluate(() => (window as any).visibilityDone)).toBe(true);
  expect(await other.evaluate(() => (window as any).visibilityObserved)).toBe('published');
  await other.close();
});

test('a later IndexedDB readwrite transaction observes a committed reference', async ({
  page,
  context,
}) => {
  const other = await context.newPage();
  await Promise.all([page.goto('/#budget'), other.goto('/#budget')]);
  await Promise.all([
    expect(page.locator('#scene-title')).toHaveText('Budget Simulator'),
    expect(other.locator('#scene-title')).toHaveText('Budget Simulator'),
  ]);
  const name = `opensheet-test-ref-${Date.now()}-${Math.random()}`;
  await page.evaluate(async (name) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const open = indexedDB.open(name, 1);
      open.onupgradeneeded = () => open.result.createObjectStore('refs');
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('refs', 'readwrite');
    const store = tx.objectStore('refs');
    let release = false;
    (window as any).releaseIdbTransaction = () => {
      release = true;
    };
    const keepAlive = () => {
      const request = store.get('ref');
      request.onsuccess = () => {
        if (release) store.put('published', 'ref');
        else keepAlive();
      };
    };
    keepAlive();
    tx.oncomplete = () => {
      (window as any).idbCommitted = true;
      db.close();
    };
    tx.onabort = () => {
      (window as any).idbError = String(tx.error);
      db.close();
    };
    (window as any).idbTransactionHeld = true;
  }, name);
  await other.evaluate(async (name) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const open = indexedDB.open(name, 1);
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('refs', 'readwrite');
    const request = tx.objectStore('refs').get('ref');
    request.onsuccess = () => {
      (window as any).idbObserved = request.result;
    };
    tx.oncomplete = () => {
      (window as any).idbReadDone = true;
      db.close();
    };
    tx.onabort = () => {
      (window as any).idbError = String(tx.error);
      db.close();
    };
    (window as any).idbReadQueued = true;
  }, name);
  await page.evaluate(() => (window as any).releaseIdbTransaction());
  await expect.poll(() => page.evaluate(() => (window as any).idbCommitted)).toBe(true);
  await expect.poll(() => other.evaluate(() => (window as any).idbReadDone)).toBe(true);
  expect(await other.evaluate(() => (window as any).idbObserved)).toBe('published');
  await other.close();
});
