import { expect, test } from '@playwright/test';

test('startup reports a changed saved scene and retries its newer reference', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const oldId = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[314]]);
    return book.id;
  });
  page.once('dialog', (dialog) => dialog.accept());
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  const newId = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[271]]);
    await book.close();
    return book.id;
  });
  const opening = await context.newPage();
  await opening.addInitScript(
    ({ next }) => {
      if (sessionStorage.getItem('reference-race-injected')) return;
      const get = Storage.prototype.getItem;
      Storage.prototype.getItem = function (key) {
        const value = get.call(this, key);
        if (this === localStorage && key === 'opensheet.scene.budget') {
          Storage.prototype.getItem = get;
          sessionStorage.setItem('reference-race-injected', 'yes');
          localStorage.setItem(key, next);
        }
        return value;
      };
    },
    { next: newId },
  );
  await opening.goto('/#budget');
  await expect(opening.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(opening.locator('#startup-recovery-message')).toContainText('Saved scene changed');
  expect(await opening.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(newId);
  await opening.getByRole('button', { name: 'Try again' }).click();
  await expect(opening.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  await expect(opening.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await opening.evaluate(() => window.opensheet.getWorkbook().id)).toBe(newId);
  expect(
    await opening.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[271]]);
  const prior = await opening.evaluate(
    async ({ id, modulePath }) => {
      const { openWorkbook } = await import(/* @vite-ignore */ modulePath);
      const book = await openWorkbook(id);
      try {
        return {
          readOnly: book.isReadOnly,
          value: await book.getSheets()[0].range('A1').getValues(),
        };
      } finally {
        await book.close();
      }
    },
    { id: oldId, modulePath: `/@fs${process.cwd()}/packages/core/src/index.ts` },
  );
  expect(prior).toEqual({ readOnly: false, value: [[314]] });
});

test('failed restore preserves the saved reference and can retry after storage is repaired', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const id = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[314]]);
    await book.close();
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const read = db.transaction('heads').objectStore('heads').get(book.id);
    (window as any).savedHead = await new Promise((resolve, reject) => {
      read.onsuccess = () => resolve(read.result);
      read.onerror = () => reject(read.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').delete(book.id);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    return book.id;
  });
  const retryPage = await context.newPage();
  await retryPage.goto('/#budget');
  await expect(retryPage.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  expect(await retryPage.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  await expect(retryPage.locator('#app')).toHaveAttribute('inert', '');
  await page.evaluate(async (id) => {
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').put((window as any).savedHead, id);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
  }, id);
  await retryPage.getByRole('button', { name: 'Try again' }).click();
  await expect(retryPage.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  await expect(retryPage.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await retryPage.evaluate(
      async () =>
        (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
    ),
  ).toBe(314);
  expect(await retryPage.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
});

test('unavailable browser storage reports a recoverable startup error', async ({ page }) => {
  await page.addInitScript(() => {
    const original = Object.getOwnPropertyDescriptor(window, 'indexedDB');
    Object.defineProperty(window, 'indexedDB', { configurable: true, value: undefined });
    (window as any).restoreIndexedDB = () =>
      original
        ? Object.defineProperty(window, 'indexedDB', original)
        : delete (window as any).indexedDB;
  });
  await page.goto('/#budget');
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('STORAGE_UNAVAILABLE');
  await expect(page.getByRole('button', { name: 'Try again' })).toBeFocused();
  await page.evaluate(() => (window as any).restoreIndexedDB());
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
});

test('denied browser storage access reports guidance and retries after access is restored', async ({
  page,
}) => {
  await page.addInitScript(() => {
    const original = Object.getOwnPropertyDescriptor(window, 'indexedDB');
    Object.defineProperty(window, 'indexedDB', {
      configurable: true,
      get() {
        throw new DOMException('Access denied', 'SecurityError');
      },
    });
    (window as any).restoreIndexedDB = () =>
      original
        ? Object.defineProperty(window, 'indexedDB', original)
        : delete (window as any).indexedDB;
  });
  await page.goto('/#budget');
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('STORAGE_UNAVAILABLE');
  await expect(page.locator('#startup-recovery-message')).toContainText('enable site storage');
  await page.evaluate(() => (window as any).restoreIndexedDB());
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
});

test('Worker database security failure preserves the saved workbook for retry', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const id = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[314]]);
    return book.id;
  });
  let inject = true;
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    const body = await response.text();
    await route.fulfill({
      response,
      body: inject
        ? `${body}\nIDBFactory.prototype.open = function() { throw new DOMException('Denied', 'SecurityError'); };`
        : body,
    });
  });
  await page.reload();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('STORAGE_UNAVAILABLE');
  await expect(page.locator('#startup-recovery-message')).toContainText('enable site storage');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  inject = false;
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  expect(
    await page.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[314]]);
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
});

test('Worker lock denial keeps the workbook intact and permits retry', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const id = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[271]]);
    return book.id;
  });
  let inject = true;
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    const body = await response.text();
    await route.fulfill({
      response,
      body: inject
        ? `${body}\nObject.getPrototypeOf(navigator.locks).request = function() { return Promise.reject(new DOMException('Denied', 'SecurityError')); };`
        : body,
    });
  });
  await page.reload();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.locator('#startup-recovery-message')).toContainText('WEB_LOCKS_UNAVAILABLE');
  await expect(page.locator('#startup-recovery-message')).toContainText('HTTPS or localhost');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  inject = false;
  await page.getByRole('button', { name: 'Try again' }).click();
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
  expect(
    await page.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[271]]);
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
});

test('failed scene restore keeps the current workbook and original saved reference', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => {
    localStorage.setItem('opensheet.scene.sales', 'missing-workbook-for-recovery-test');
    return window.opensheet.getWorkbook().id;
  });
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#notice')).toContainText('Workbook does not exist');
  await expect(page.locator('#scene-select')).toHaveValue('budget');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.sales'))).toBe(
    'missing-workbook-for-recovery-test',
  );
  expect(new URL(page.url()).hash).toBe('#budget');
});

test('failed reset keeps the saved workbook reference', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const original = await page.evaluate(() => {
    const originalSet = Storage.prototype.setItem;
    Storage.prototype.setItem = function (key, value) {
      if (key === 'opensheet.scene.budget') throw new Error('Injected storage failure');
      return originalSet.call(this, key, value);
    };
    (window as any).restoreSetItem = () => (Storage.prototype.setItem = originalSet);
    return window.opensheet.getWorkbook().id;
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.locator('#notice')).toHaveText('Injected storage failure');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(original);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
  await page.evaluate(() => (window as any).restoreSetItem());
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(original);
});

test('reset publishes a new workbook and retains the previous workbook', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const oldId = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[314]]);
    return book.id;
  });
  page.once('dialog', (dialog) => dialog.accept());
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect
    .poll(() => page.evaluate(() => localStorage.getItem('opensheet.scene.budget')))
    .not.toBe(oldId);
  const newId = await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'));
  await expect.poll(() => page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(newId);
  const other = await context.newPage();
  await other.goto('/#sales');
  await expect(other.locator('#scene-title')).toHaveText('Sales Dashboard');
  await other.evaluate((id) => window.opensheet.open(id), oldId);
  expect(
    await other.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[314]]);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(newId);
});

test('reset retains the published workbook if attachment fails', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const oldId = await page.evaluate(() => {
    window.opensheet.attachWorkbook = async () => {
      throw new Error('Injected attachment failure');
    };
    return window.opensheet.getWorkbook().id;
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect(page.locator('#notice')).toContainText('Injected attachment failure');
  await expect(page.locator('#notice')).toContainText('was retained');
  const newId = await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'));
  expect(newId).toBeTruthy();
  expect(newId).not.toBe(oldId);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(oldId);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(newId);
});

test('reset aborts if the original workbook changes during preparation', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const oldId = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const prototype = Object.getPrototypeOf(book);
    const ready = prototype.ready;
    prototype.ready = async function () {
      const result = await ready.call(this);
      if (this !== book) {
        (window as any).resetCandidate = this.id;
        await new Promise<void>((resolve) => ((window as any).releaseReset = resolve));
      }
      return result;
    };
    return book.id;
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect.poll(() => page.evaluate(() => Boolean((window as any).releaseReset))).toBe(true);
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[628]]);
    (window as any).releaseReset();
  });
  await expect(page.locator('#notice')).toContainText('Workbook changed during reset');
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(oldId);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(oldId);
  const deleted = await page.evaluate(async () => {
    const request = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const read = db
      .transaction('heads')
      .objectStore('heads')
      .get((window as any).resetCandidate);
    const result = await new Promise((resolve, reject) => {
      read.onsuccess = () => resolve(read.result);
      read.onerror = () => reject(read.error);
    });
    db.close();
    return result === undefined;
  });
  expect(deleted).toBe(true);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await page.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[628]]);
});

test('scene navigation during reset keeps the address and workbook consistent', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const prototype = Object.getPrototypeOf(book);
    const ready = prototype.ready;
    prototype.ready = async function () {
      const result = await ready.call(this);
      if (this !== book) {
        prototype.ready = ready;
        await new Promise<void>((resolve) => ((window as any).releaseReset = resolve));
      }
      return result;
    };
  });
  await page.getByRole('button', { name: 'Reset demo' }).click();
  await expect.poll(() => page.evaluate(() => Boolean((window as any).releaseReset))).toBe(true);
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#notice')).toContainText('Wait for the current scene operation');
  await expect(page.locator('#scene-select')).toHaveValue('budget');
  expect(new URL(page.url()).hash).toBe('#budget');
  await page.evaluate(() => (window as any).releaseReset());
  await expect(page.getByRole('button', { name: 'Reset demo' })).toBeEnabled();
  const id = await page.evaluate(() => window.opensheet.getWorkbook().id);
  expect(await page.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(id);
  await page.reload();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
});

for (const format of ['json', 'opensheet'] as const) {
  test(`restores a ${format} backup only after a complete import`, async ({ page, context }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const fixture = await page.evaluate(async (format) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[9876]]);
      const chunks: number[] = [];
      if (format === 'json') {
        const text: string[] = [];
        for await (const part of book.streamJSON()) text.push(part);
        chunks.push(...new TextEncoder().encode(text.join('')));
      } else {
        for await (const part of book.streamBinary()) chunks.push(...part);
      }
      await book.close();
      const open = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
      });
      const tx = db.transaction('heads', 'readwrite');
      tx.objectStore('heads').delete(book.id);
      await new Promise<void>((resolve, reject) => {
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id: book.id, chunks };
    }, format);
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    const input = recovery.locator('#restore-backup-file');
    const heads = () =>
      recovery.evaluate(async () => {
        const open = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          open.onsuccess = () => resolve(open.result);
          open.onerror = () => reject(open.error);
        });
        const request = db.transaction('heads').objectStore('heads').getAllKeys();
        const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        db.close();
        return keys.map(String).sort();
      });
    const headsBefore = await heads();
    const verifyInput = recovery.locator('#verify-backup-file');
    const chooserPromise = recovery.waitForEvent('filechooser');
    await recovery.getByRole('button', { name: 'Verify backup' }).click();
    const chooser = await chooserPromise;
    await chooser.setFiles({
      name: `invalid.${format}`,
      mimeType: 'application/octet-stream',
      buffer: Buffer.from('invalid data'),
    });
    await expect(recovery.locator('#verify-backup-status')).not.toBeEmpty();
    await expect(recovery.getByRole('button', { name: 'Verify backup' })).toBeEnabled();
    expect(await heads()).toEqual(headsBefore);
    if (format === 'json') {
      await recovery.evaluate(() => {
        const original = window.Worker;
        (window as any).restoreVerificationWorker = () => {
          window.Worker = original;
        };
        window.Worker = new Proxy(original, {
          construct(target, args) {
            if (!String(args[0]).includes('backup-verifier'))
              return Reflect.construct(target, args);
            return {
              postMessage() {
                (window as any).verificationStarted = true;
              },
              terminate() {
                (window as any).verificationTerminated = true;
              },
            };
          },
        });
      });
      await verifyInput.setInputFiles({
        name: 'delayed-backup.json',
        mimeType: 'application/json',
        buffer: Buffer.from(fixture.chunks),
      });
      await expect
        .poll(() => recovery.evaluate(() => (window as any).verificationStarted))
        .toBe(true);
      const cancel = recovery.getByRole('button', { name: 'Cancel verification' });
      await expect(cancel).toBeVisible();
      await cancel.click();
      await expect(recovery.locator('#verify-backup-status')).toHaveText('Verification cancelled.');
      expect(await recovery.evaluate(() => (window as any).verificationTerminated)).toBe(true);
      await expect(cancel).toBeHidden();
      await expect(recovery.getByRole('button', { name: 'Verify backup' })).toBeFocused();
      await recovery.evaluate(() => (window as any).restoreVerificationWorker());
      expect(await heads()).toEqual(headsBefore);
      expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
        fixture.id,
      );
    }
    await verifyInput.setInputFiles({
      name: `backup.${format}`,
      mimeType: 'application/octet-stream',
      buffer: Buffer.from(fixture.chunks),
    });
    await expect(recovery.locator('#verify-backup-status')).toContainText('Structure verified:');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
      fixture.id,
    );
    expect(await heads()).toEqual(headsBefore);
    await input.setInputFiles({
      name: `invalid.${format}`,
      mimeType: 'application/octet-stream',
      buffer: Buffer.from('invalid data'),
    });
    await expect(recovery.locator('#startup-recovery-message')).not.toBeEmpty();
    await expect(recovery.getByRole('button', { name: 'Restore backup' })).toBeEnabled();
    expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
      fixture.id,
    );
    expect(await heads()).toEqual(headsBefore);
    if (format === 'json') {
      await recovery.evaluate(() => {
        const original = Storage.prototype.setItem;
        (window as any).restoreSetItem = () => {
          Storage.prototype.setItem = original;
        };
        Storage.prototype.setItem = function (key, value) {
          if (key === 'opensheet.scene.budget') throw new Error('Injected reference failure');
          return original.call(this, key, value);
        };
      });
      await input.setInputFiles({
        name: 'backup.json',
        mimeType: 'application/json',
        buffer: Buffer.from(fixture.chunks),
      });
      await expect(recovery.locator('#startup-recovery-message')).toHaveText(
        'Injected reference failure',
      );
      await expect(
        recovery.getByRole('heading', { name: 'Unable to open workbook' }),
      ).toBeVisible();
      expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
        fixture.id,
      );
      expect(await heads()).toEqual(headsBefore);
      await recovery.evaluate(() => (window as any).restoreSetItem());
    }
    await input.setInputFiles({
      name: `backup.${format}`,
      mimeType: 'application/octet-stream',
      buffer: Buffer.from(fixture.chunks),
    });
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
    await expect(recovery.locator('#scene-title')).toHaveText('Budget Simulator');
    const restored = await recovery.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      value: (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
      saved: localStorage.getItem('opensheet.scene.budget'),
    }));
    expect(restored.id).not.toBe(fixture.id);
    expect(restored.saved).toBe(restored.id);
    expect(restored.value).toBe(9876);
    await recovery.reload();
    await expect(recovery.locator('#scene-title')).toHaveText('Budget Simulator');
    expect(
      await recovery.evaluate(
        async () =>
          (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
      ),
    ).toBe(9876);
    await recovery.close();
  });
}

test('performance startup failure offers a safe route to a regular workbook', async ({ page }) => {
  await page.addInitScript(() => {
    if (location.hash !== '#performance') return;
    Object.defineProperty(window, 'indexedDB', { configurable: true, value: undefined });
  });
  await page.goto('/#performance');
  await expect(page.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(page.getByRole('button', { name: 'Restore backup' })).toBeHidden();
  await page.getByRole('button', { name: 'Open Budget Simulator' }).click();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(new URL(page.url()).hash).toBe('#budget');
});

test('performance storage corruption does not trigger destructive initialization', async ({
  page,
  context,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book.close();
    const open = indexedDB.open('opensheet-performance');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').put(new Uint8Array([9]), 'performance-workbook');
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
  });
  const other = await context.newPage();
  await other.goto('/#performance');
  await expect(other.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(other.locator('#startup-recovery-message')).toContainText('CORRUPT_STORAGE');
  expect(
    await other.evaluate(async () => {
      const open = indexedDB.open('opensheet-performance');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        open.onsuccess = () => resolve(open.result);
        open.onerror = () => reject(open.error);
      });
      const request = db.transaction('heads').objectStore('heads').get('performance-workbook');
      const value = await new Promise<Uint8Array>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      db.close();
      return Array.from(value);
    }),
  ).toEqual([9]);
  await other.close();
});

test('a missing performance head does not delete remaining workbook records', async ({
  page,
  context,
}) => {
  await page.goto('/#performance');
  await expect(page.locator('#scene-title')).toHaveText('Performance Lab');
  const before = await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[2468]]);
    await window.opensheet.getWorkbook().close();
    const open = indexedDB.open('opensheet-performance');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').delete('performance-workbook');
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    const get = db.transaction('records').objectStore('records').getAllKeys();
    const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
      get.onsuccess = () => resolve(get.result);
      get.onerror = () => reject(get.error);
    });
    db.close();
    return keys.map(String).sort();
  });
  expect(before.length).toBeGreaterThan(0);
  const other = await context.newPage();
  await other.goto('/#performance');
  await expect(other.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await expect(other.locator('#startup-recovery-message')).toContainText('CORRUPT_STORAGE');
  const after = await other.evaluate(async () => {
    const open = indexedDB.open('opensheet-performance');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const get = db.transaction('records').objectStore('records').getAllKeys();
    const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
      get.onsuccess = () => resolve(get.result);
      get.onerror = () => reject(get.error);
    });
    db.close();
    return keys.map(String).sort();
  });
  expect(after).toEqual(before);
  await other.close();
});

test('core defaults preserve records when creating with a missing head', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const existing = window.opensheet.getWorkbook();
    const Constructor = existing.constructor as new (
      snapshot: any,
      options: object,
    ) => typeof existing;
    const snapshot = await existing.toJSON();
    snapshot.workbookId = `missing-${crypto.randomUUID()}`;
    const database = `missing-test-${crypto.randomUUID()}`;
    const original = await new Constructor(snapshot, { database }).ready();
    await original.close();
    const open = indexedDB.open(database);
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').delete(snapshot.workbookId);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    const keys = async () => {
      const get = db.transaction('records').objectStore('records').getAllKeys();
      return (
        await new Promise<IDBValidKey[]>((resolve, reject) => {
          get.onsuccess = () => resolve(get.result);
          get.onerror = () => reject(get.error);
        })
      )
        .map(String)
        .sort();
    };
    const before = await keys();
    let code = 'unexpected success';
    try {
      await new Constructor(snapshot, { database }).ready();
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    const after = await keys();
    db.close();
    return { code, before, after };
  });
  expect(result.code).toBe('CORRUPT_STORAGE');
  expect(result.before.length).toBeGreaterThan(0);
  expect(result.after).toEqual(result.before);
});

for (const failure of ['attachment'] as const) {
  test(`preserves an attached restored workbook after a ${failure} failure`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const fixture = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[9876]]);
      const chunks: string[] = [];
      for await (const part of book.streamJSON()) chunks.push(part);
      await book.close();
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('opensheet');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction('heads', 'readwrite');
        tx.objectStore('heads').delete(book.id);
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id: book.id, json: chunks.join('') };
    });
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await recovery.evaluate(() => {
      const app = window.opensheet;
      const attach = app.attachWorkbook.bind(app);
      app.attachWorkbook = async (book) => {
        await attach(book);
        throw new Error('Injected attachment failure');
      };
    });
    await recovery.locator('#restore-backup-file').setInputFiles({
      name: 'backup.json',
      mimeType: 'application/json',
      buffer: Buffer.from(fixture.json),
    });
    await expect(recovery.locator('#startup-recovery-message')).toContainText(
      `Injected ${failure} failure`,
    );
    const result = await recovery.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const id = book.id;
      const value = await book.getSheets()[0].range('A1').getValues();
      await window.opensheet.open(id);
      return {
        id,
        value,
        reopened: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
        reference: localStorage.getItem('opensheet.scene.budget'),
      };
    });
    expect(result.id).not.toBe(fixture.id);
    await expect(recovery.locator('#startup-recovery-message')).toContainText(result.id);
    await expect(recovery.locator('#startup-recovery-message')).toContainText('retained');
    expect(result.value).toEqual([[9876]]);
    expect(result.reopened).toEqual([[9876]]);
    expect(result.reference).toBe(result.id);
    await recovery.close();
  });
}

for (const failure of ['before-attach', 'rollback-trap'] as const) {
  test(`recovery handles ${failure} failure without a dangling scene reference`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const fixture = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[9876]]);
      const parts: string[] = [];
      for await (const part of book.streamJSON()) parts.push(part);
      await book.close();
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('opensheet');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction('heads', 'readwrite');
        tx.objectStore('heads').delete(book.id);
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id: book.id, json: parts.join('') };
    });
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    const heads = () =>
      recovery.evaluate(async () => {
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          const request = indexedDB.open('opensheet');
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const keys = await new Promise<IDBValidKey[]>((resolve, reject) => {
          const request = db.transaction('heads').objectStore('heads').getAllKeys();
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        db.close();
        return keys.map(String).sort();
      });
    await recovery.evaluate(
      ({ failure, oldId }) => {
        const app = window.opensheet;
        const originalAttach = app.attachWorkbook.bind(app);
        (window as any).restoreRecoveryAttach = () => {
          app.attachWorkbook = originalAttach;
        };
        app.attachWorkbook = async () => {
          throw new Error('Injected pre-attachment failure');
        };
        if (failure === 'rollback-trap') {
          const set = Storage.prototype.setItem;
          (window as any).restoreSetItem = () => {
            Storage.prototype.setItem = set;
          };
          Storage.prototype.setItem = function (key, value) {
            if (key === 'opensheet.scene.budget' && value === oldId) {
              (window as any).rollbackAttempted = true;
              throw new Error('Injected rollback failure');
            }
            return set.call(this, key, value);
          };
        }
      },
      { failure, oldId: fixture.id },
    );
    await recovery.locator('#restore-backup-file').setInputFiles({
      name: 'backup.json',
      mimeType: 'application/json',
      buffer: Buffer.from(fixture.json),
    });
    await expect(recovery.locator('#startup-recovery-message')).toContainText(
      'Injected pre-attachment failure',
    );
    const reference = await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'));
    expect(reference).not.toBeNull();
    expect(reference).not.toBe(fixture.id);
    await expect(recovery.locator('#startup-recovery-message')).toContainText(reference!);
    await expect(recovery.locator('#startup-recovery-message')).toContainText('retained');
    expect(await heads()).toContain(reference);
    if (failure === 'rollback-trap')
      expect(await recovery.evaluate(() => (window as any).rollbackAttempted)).toBeUndefined();
    await recovery.evaluate(() => (window as any).restoreSetItem?.());
    if (failure === 'rollback-trap') {
      await page.evaluate((id) => localStorage.setItem('opensheet.scene.budget', id), fixture.id);
      await expect
        .poll(() => recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget')))
        .toBe(fixture.id);
      const previousId = await recovery.evaluate(async (json) => {
        (window as any).restoreRecoveryAttach();
        const app = window.opensheet;
        await app.load(JSON.parse(json));
        const previous = app.getWorkbook();
        await previous
          .getSheets()[0]
          .range('A1')
          .setValues([[4321]]);
        return previous.id;
      }, fixture.json);
      await recovery.evaluate(() => {
        const original = Storage.prototype.setItem;
        (window as any).restoreRetainedSetItem = () => {
          Storage.prototype.setItem = original;
        };
        Storage.prototype.setItem = function (key, value) {
          if (key === 'opensheet.scene.budget')
            throw new Error('Injected retained reference failure');
          return original.call(this, key, value);
        };
      });
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
      await expect(recovery.locator('#startup-recovery-message')).toContainText(
        'Injected retained reference failure',
      );
      await expect(recovery.getByRole('button', { name: 'Open retained backup' })).toBeEnabled();
      expect(
        await recovery.evaluate(async () => {
          const book = window.opensheet.getWorkbook();
          return { id: book.id, value: await book.getSheets()[0].range('A1').getValues() };
        }),
      ).toEqual({ id: previousId, value: [[4321]] });
      expect(
        await recovery.evaluate(
          async (id) => ({
            reference: localStorage.getItem('opensheet.scene.budget'),
            candidateLocked: ((await navigator.locks.query()).held ?? []).some(
              (lock) => lock.name === `opensheet:opensheet:${id}`,
            ),
          }),
          reference,
        ),
      ).toEqual({ reference: fixture.id, candidateLocked: false });
      await recovery.evaluate(() => {
        (window as any).restoreRetainedSetItem();
        (window as any).restoreRecoveryAttach();
        const original = Storage.prototype.getItem;
        let reads = 0;
        Storage.prototype.getItem = function (key) {
          if (key === 'opensheet.scene.budget' && ++reads === 2) {
            Storage.prototype.getItem = original;
            this.setItem(key, 'other-scene-reference');
          }
          return original.call(this, key);
        };
      });
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
      await expect(recovery.locator('#startup-recovery-message')).toContainText(
        'saved scene changed',
      );
      expect(
        await recovery.evaluate(() => ({
          id: window.opensheet.getWorkbook().id,
          reference: localStorage.getItem('opensheet.scene.budget'),
        })),
      ).toEqual({ id: previousId, reference: 'other-scene-reference' });
      await recovery.evaluate((id) => {
        localStorage.setItem('opensheet.scene.budget', id);
        window.opensheet.attachWorkbook = async () => {
          throw new Error('Injected retained attachment failure');
        };
      }, fixture.id);
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
      await expect(recovery.locator('#startup-recovery-message')).toContainText(
        'Injected retained attachment failure',
      );
      expect(
        await recovery.evaluate(
          async (id) => ({
            current: window.opensheet.getWorkbook().id,
            reference: localStorage.getItem('opensheet.scene.budget'),
            candidateLocked: ((await navigator.locks.query()).held ?? []).some(
              (lock) => lock.name === `opensheet:opensheet:${id}`,
            ),
          }),
          reference,
        ),
      ).toEqual({ current: previousId, reference, candidateLocked: false });
      await recovery.evaluate(() => {
        (window as any).restoreRecoveryAttach();
        const attach = window.opensheet.attachWorkbook.bind(window.opensheet);
        window.opensheet.attachWorkbook = async (book) => {
          await attach(book);
          throw new Error('Injected failure after retained attachment');
        };
      });
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
      await expect(recovery.locator('#startup-recovery-message')).toContainText(
        'Injected failure after retained attachment',
      );
      expect(
        await recovery.evaluate(() => ({
          current: window.opensheet.getWorkbook().id,
          reference: localStorage.getItem('opensheet.scene.budget'),
        })),
      ).toEqual({ current: reference, reference });
      await recovery.evaluate(() => (window as any).restoreRecoveryAttach());
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
    } else {
      await recovery.evaluate(() => (window as any).restoreRecoveryAttach());
      await recovery.getByRole('button', { name: 'Try again' }).click();
    }
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
    const result = await recovery.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      saved: localStorage.getItem('opensheet.scene.budget'),
    }));
    expect(result).toEqual({ id: reference, value: [[9876]], saved: reference });
    await recovery.reload();
    await expect(recovery.locator('#scene-title')).toHaveText('Budget Simulator');
    expect(await recovery.evaluate(() => window.opensheet.getWorkbook().id)).toBe(reference);
    await recovery.close();
  });
}

for (const phase of ['before', 'after'] as const) {
  test(`reopens a complete backup after reload ${phase} attachment completes`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const fixture = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[9876]]);
      const parts: string[] = [];
      for await (const part of book.streamJSON()) parts.push(part);
      await book.close();
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        const request = indexedDB.open('opensheet');
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction('heads', 'readwrite');
        tx.objectStore('heads').delete(book.id);
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id: book.id, json: parts.join('') };
    });
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    await recovery.evaluate((phase) => {
      const app = window.opensheet;
      const attach = app.attachWorkbook.bind(app);
      app.attachWorkbook = async (book) => {
        if (phase === 'after') await attach(book);
        (window as any).attachReached = true;
        await new Promise(() => {});
      };
    }, phase);
    await recovery.locator('#restore-backup-file').setInputFiles({
      name: 'backup.json',
      mimeType: 'application/json',
      buffer: Buffer.from(fixture.json),
    });
    await expect.poll(() => recovery.evaluate(() => (window as any).attachReached)).toBe(true);
    const published = await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'));
    expect(published).not.toBeNull();
    expect(published).not.toBe(fixture.id);
    await recovery.reload();
    await expect(recovery.locator('#scene-select')).toBeEnabled();
    const result = await recovery.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    }));
    expect(result).toEqual({ id: published, value: [[9876]] });
    await recovery.close();
  });
}

test('failed backup attachment does not overwrite another tab scene reference', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const fixture = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[9876]]);
    const snapshot = await book.toJSON();
    const parts: string[] = [];
    for await (const part of book.streamJSON()) parts.push(part);
    const Workbook = book.constructor as new (file: typeof snapshot) => typeof book;
    const other = new Workbook({ ...snapshot, workbookId: `other_${crypto.randomUUID()}` });
    await other.ready();
    await other
      .getSheets()[0]
      .range('A1')
      .setValues([[123]]);
    await other.close();
    await book.close();
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('opensheet');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('heads', 'readwrite');
      tx.objectStore('heads').delete(book.id);
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
    return { id: book.id, otherId: other.id, json: parts.join('') };
  });
  const recovery = await context.newPage();
  await recovery.goto('/#budget');
  await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
  await recovery.evaluate(() => {
    let release!: () => void;
    const pending = new Promise<void>((resolve) => {
      release = resolve;
    });
    (window as any).releaseAttach = release;
    window.opensheet.attachWorkbook = async () => {
      (window as any).attachReached = true;
      await pending;
      throw new Error('Injected pre-attachment failure');
    };
  });
  await recovery.locator('#restore-backup-file').setInputFiles({
    name: 'backup.json',
    mimeType: 'application/json',
    buffer: Buffer.from(fixture.json),
  });
  await expect.poll(() => recovery.evaluate(() => (window as any).attachReached)).toBe(true);
  const candidate = await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'));
  expect(candidate).not.toBe(fixture.id);
  expect(candidate).not.toBe(fixture.otherId);
  await page.evaluate(async (id) => {
    await window.opensheet.open(id!);
  }, candidate);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[9876]]);
  await page.evaluate((id) => localStorage.setItem('opensheet.scene.budget', id), fixture.otherId);
  await recovery.evaluate(() => (window as any).releaseAttach());
  await expect(recovery.locator('#startup-recovery-message')).toContainText(
    'Injected pre-attachment failure',
  );
  expect(await recovery.evaluate(() => localStorage.getItem('opensheet.scene.budget'))).toBe(
    fixture.otherId,
  );
  const hasCandidate = await recovery.evaluate(async (id) => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('opensheet');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const exists = await new Promise<boolean>((resolve, reject) => {
      const request = db.transaction('heads').objectStore('heads').get(id!);
      request.onsuccess = () => resolve(request.result !== undefined);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return exists;
  }, candidate);
  expect(hasCandidate).toBe(true);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[9876]]);
  await expect(recovery.locator('#startup-recovery-message')).toContainText(candidate!);
  const retainedValue = await recovery.evaluate(async (id) => {
    await window.opensheet.open(id!);
    const book = window.opensheet.getWorkbook();
    const value = await book.getSheets()[0].range('A1').getValues();
    await book.close();
    return value;
  }, candidate);
  expect(retainedValue).toEqual([[9876]]);
  await recovery.reload();
  await expect(recovery.locator('#scene-select')).toBeEnabled();
  expect(
    await recovery.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    })),
  ).toEqual({ id: fixture.otherId, value: [[123]] });
  await recovery.close();
});
