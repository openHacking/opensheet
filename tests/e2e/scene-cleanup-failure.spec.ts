import { expect, test } from '@playwright/test';

for (const operation of ['creation', 'reset', 'published'] as const) {
  test(`${operation} preserves its original error when candidate cleanup also throws`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async (operation) => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[731]]);
      const prototype = Object.getPrototypeOf(book);
      const original = {
        dispose: prototype.dispose,
        deleteStorage: prototype.deleteStorage,
        close: prototype.close,
        setItem: Storage.prototype.setItem,
        attach: app.attachWorkbook,
      };
      (window as any).restoreCandidateHooks = () => {
        prototype.dispose = original.dispose;
        prototype.deleteStorage = original.deleteStorage;
        prototype.close = original.close;
        Storage.prototype.setItem = original.setItem;
        app.attachWorkbook = original.attach;
      };
      prototype.dispose = function () {
        if (this.id !== (window as any).cleanupCandidate) return original.dispose.call(this);
        original.dispose.call(this);
        throw new Error('Injected dispose failure');
      };
      prototype.deleteStorage = async function () {
        if (this.id !== (window as any).cleanupCandidate) return original.deleteStorage.call(this);
        throw new Error('Injected delete failure');
      };
      prototype.close = async function () {
        if (this.id !== (window as any).cleanupCandidate) return original.close.call(this);
        await original.close.call(this).catch(() => {});
        throw new Error('Injected close failure');
      };
      Storage.prototype.setItem = function (key, value) {
        if (
          this === localStorage &&
          key === (operation === 'reset' ? 'opensheet.scene.budget' : 'opensheet.scene.sales')
        ) {
          (window as any).cleanupCandidate = value;
          if (operation !== 'published') throw new Error('Injected reference failure');
        }
        return original.setItem.call(this, key, value);
      };
      if (operation === 'published')
        app.attachWorkbook = async () => {
          throw new Error('Injected attach failure');
        };
      return {
        id: book.id,
        revision: book.revision,
        reference: localStorage.getItem(
          operation === 'reset' ? 'opensheet.scene.budget' : 'opensheet.scene.sales',
        ),
      };
    }, operation);
    page.once('dialog', (dialog) => dialog.accept());
    if (operation === 'reset') await page.locator('#reset-demo').click();
    else await page.locator('#scene-select').selectOption('sales');
    const notice = page.locator('#notice');
    await expect(notice).toContainText(
      operation === 'published' ? 'Injected attach failure' : 'Injected reference failure',
    );
    await expect(notice).not.toContainText('Injected dispose failure');
    const actual = await page.evaluate(async (operation) => {
      (window as any).restoreCandidateHooks();
      const book = window.opensheet.getWorkbook();
      const candidate = (window as any).cleanupCandidate as string;
      const reference = localStorage.getItem(
        operation === 'reset' ? 'opensheet.scene.budget' : 'opensheet.scene.sales',
      );
      const path = `/@fs${location.pathname}`; // The caller supplies the module URL below instead.
      void path;
      return {
        id: book.id,
        revision: book.revision,
        value: await book.getSheets()[0].range('A1').getValues(),
        candidate,
        reference,
      };
    }, operation);
    expect(actual).toMatchObject({
      id: before.id,
      revision: before.revision,
      value: [[731]],
      candidate: expect.any(String),
      reference: operation === 'published' ? actual.candidate : before.reference,
    });
    const recovered = await page.evaluate(
      async ({ url, id }) => {
        const { openWorkbook } = await import(/* @vite-ignore */ url);
        const candidate = await openWorkbook(id);
        const readOnly = candidate.isReadOnly;
        await candidate.deleteStorage();
        return readOnly;
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, id: actual.candidate },
    );
    expect(recovered).toBe(false);
  });
}

for (const phase of ['unpublished', 'published'] as const) {
  test(`backup ${phase} failure survives throwing candidate cleanup`, async ({ page, context }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const fixture = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[8642]]);
      const json = await book.toJSON();
      const id = book.id;
      await book.close();
      const request = indexedDB.open('opensheet');
      const db = await new Promise<IDBDatabase>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const tx = db.transaction('heads', 'readwrite');
      tx.objectStore('heads').delete(id);
      await new Promise<void>((resolve, reject) => {
        tx.oncomplete = () => resolve();
        tx.onabort = () => reject(tx.error);
      });
      db.close();
      return { id, json: JSON.stringify(json) };
    });
    const recovery = await context.newPage();
    await recovery.goto('/#budget');
    await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeVisible();
    const previous = await recovery.evaluate(
      async ({ json, phase }) => {
        const app = window.opensheet;
        await app.load(JSON.parse(json));
        const book = app.getWorkbook();
        const proto = Object.getPrototypeOf(book);
        const original = {
          dispose: proto.dispose,
          deleteStorage: proto.deleteStorage,
          close: proto.close,
          importJSON: proto.importJSON,
          setItem: Storage.prototype.setItem,
          attach: app.attachWorkbook,
        };
        (window as any).restoreBackupHooks = () => {
          Object.assign(proto, {
            dispose: original.dispose,
            deleteStorage: original.deleteStorage,
            close: original.close,
            importJSON: original.importJSON,
          });
          Storage.prototype.setItem = original.setItem;
          app.attachWorkbook = original.attach;
        };
        proto.importJSON = async function (...args: unknown[]) {
          const result = await original.importJSON.apply(this, args);
          if (this !== book) (window as any).cleanupBackup = this.id;
          return result;
        };
        proto.dispose = function () {
          if (this.id !== (window as any).cleanupBackup) return original.dispose.call(this);
          original.dispose.call(this);
          throw new Error('Injected backup dispose failure');
        };
        proto.deleteStorage = async function () {
          if (this.id !== (window as any).cleanupBackup) return original.deleteStorage.call(this);
          throw new Error('Injected backup delete failure');
        };
        proto.close = async function () {
          if (this.id !== (window as any).cleanupBackup) return original.close.call(this);
          await original.close.call(this).catch(() => {});
          throw new Error('Injected backup close failure');
        };
        Storage.prototype.setItem = function (key, value) {
          if (phase === 'unpublished' && this === localStorage && key === 'opensheet.scene.budget')
            throw new Error('Injected backup reference failure');
          return original.setItem.call(this, key, value);
        };
        if (phase === 'published')
          app.attachWorkbook = async () => {
            throw new Error('Injected backup attachment failure');
          };
        return { id: book.id, revision: book.revision };
      },
      { json: fixture.json, phase },
    );
    await recovery.locator('#restore-backup-file').setInputFiles({
      name: 'backup.json',
      mimeType: 'application/json',
      buffer: Buffer.from(fixture.json),
    });
    const message = recovery.locator('#startup-recovery-message');
    await expect(message).toContainText(
      phase === 'published'
        ? 'Injected backup attachment failure'
        : 'Injected backup reference failure',
    );
    await expect(message).not.toContainText('Injected backup dispose failure');
    if (phase === 'published')
      await expect(recovery.getByRole('button', { name: 'Open retained backup' })).toBeVisible();
    const state = await recovery.evaluate(async () => {
      (window as any).restoreBackupHooks();
      const book = window.opensheet.getWorkbook();
      return {
        id: book.id,
        revision: book.revision,
        value: await book.getSheets()[0].range('A1').getValues(),
        candidate: (window as any).cleanupBackup as string,
        reference: localStorage.getItem('opensheet.scene.budget'),
      };
    });
    expect(state).toMatchObject({
      ...previous,
      value: [[8642]],
      candidate: expect.any(String),
      reference: phase === 'published' ? state.candidate : fixture.id,
    });
    if (phase === 'published') {
      await recovery.evaluate((id) => {
        const proto = Object.getPrototypeOf(window.opensheet.getWorkbook());
        const close = proto.close;
        const dispose = proto.dispose;
        const setItem = Storage.prototype.setItem;
        (window as any).restoreRetainedHooks = () => {
          proto.close = close;
          proto.dispose = dispose;
          Storage.prototype.setItem = setItem;
        };
        proto.close = async function () {
          if (this.id !== id) return close.call(this);
          await close.call(this).catch(() => {});
          throw new Error('Injected retained close failure');
        };
        proto.dispose = function () {
          if (this.id !== id) return dispose.call(this);
          dispose.call(this);
          throw new Error('Injected retained dispose failure');
        };
        Storage.prototype.setItem = function (key, value) {
          if (this === localStorage && key === 'opensheet.scene.budget')
            throw new Error('Injected retained reference failure');
          return setItem.call(this, key, value);
        };
      }, state.candidate);
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
      await expect(message).toContainText('Injected retained reference failure');
      await expect(recovery.getByRole('button', { name: 'Open retained backup' })).toBeEnabled();
      await recovery.evaluate(() => (window as any).restoreRetainedHooks());
      await recovery.getByRole('button', { name: 'Open retained backup' }).click();
      await expect(recovery.getByRole('heading', { name: 'Unable to open workbook' })).toBeHidden();
      expect(await recovery.evaluate(() => window.opensheet.getWorkbook().id)).toBe(
        state.candidate,
      );
    } else {
      const unlocked = await recovery.evaluate(
        async ({ url, id }) => {
          const { openWorkbook } = await import(/* @vite-ignore */ url);
          const candidate = await openWorkbook(id);
          const writable = !candidate.isReadOnly;
          await candidate.deleteStorage();
          return writable;
        },
        { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, id: state.candidate },
      );
      expect(unlocked).toBe(true);
    }
  });
}
