import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
});

test('saving remains visible until every durable acknowledgement is received', async ({ page }) => {
  const snapshot = await page.evaluate(async () =>
    JSON.stringify(await window.opensheet.getWorkbook().toJSON()),
  );
  const before = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    const ids = new Set<number>();
    const held: MessageEvent[] = [];
    (window as any).heldWrites = held;
    (window as any).releaseWrite = (index: number) => handler.call(worker, held[index]);
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') ids.add(message.id);
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (ids.has(event.data.id)) held.push(event);
      else handler.call(worker, event);
    };
    const sheet = book.getSheets()[0];
    (window as any).writes = Promise.all([
      sheet.range('B2').setValues([[11]]),
      sheet.range('B3').setValues([[12]]),
    ]);
    return book.revision;
  });
  await expect(page.locator('#save-state')).toHaveText('Saving…');
  await page.waitForFunction(() => (window as any).heldWrites.length === 2);
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  expect(
    await page.evaluate(() => {
      const event = new Event('beforeunload', { cancelable: true });
      window.dispatchEvent(event);
      return event.defaultPrevented;
    }),
  ).toBe(true);
  expect(
    await page.evaluate(async (file) => {
      try {
        await window.opensheet.load(JSON.parse(file));
        return 'loaded';
      } catch (error) {
        return (error as { code?: string }).code;
      }
    }, snapshot),
  ).toBe('BUSY');
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#scene-select')).toHaveValue('budget');
  await expect(page.locator('#notice')).toContainText('Wait for saving to finish');
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await expect(page.locator('#notice')).toContainText('Wait for saving to finish');
  await expect(page.locator('#save-state')).toHaveText('Saving…');
  await page.evaluate(() => (window as any).releaseWrite(0));
  await expect(page.locator('#save-state')).toHaveText('Saving…');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending)).toBe(
    1,
  );
  await page.evaluate(async () => {
    (window as any).releaseWrite(1);
    await (window as any).writes;
  });
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending)).toBe(
    0,
  );
  const reopened = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const Constructor = book.constructor as new (
      snapshot: undefined,
      options: object,
      id: string,
    ) => typeof book;
    const reader = await new Constructor(undefined, {}, book.id).ready();
    try {
      return await reader.getSheets()[0].range('B2:B3').getValues();
    } finally {
      await reader.close();
    }
  });
  expect(reopened).toEqual([[11], [12]]);
});

test('a transaction is saving while its asynchronous work is still pending', async ({ page }) => {
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    (window as any).releaseTransaction = release;
    (window as any).transaction = book.transaction({}, async (tx) => {
      await gate;
      await tx
        .getSheets()[0]
        .range('B2')
        .setValues([[23]]);
    });
  });
  await expect(page.locator('#save-state')).toHaveText('Saving…');
  await page.evaluate(async () => {
    (window as any).releaseTransaction();
    await (window as any).transaction;
  });
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
});

test('a write starting during replacement prevents disposal of the original workbook', async ({
  page,
}) => {
  const initialId = await page.evaluate(async () => {
    const app = window.opensheet;
    const old = app.getWorkbook();
    const snapshot = await old.toJSON();
    const oldWorker = (old as any).worker as Worker;
    const post = oldWorker.postMessage.bind(oldWorker);
    const handler = oldWorker.onmessage!;
    let writeId = -1;
    oldWorker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') writeId = message.id;
      return Reflect.apply(post, oldWorker, [message, ...rest]);
    }) as typeof oldWorker.postMessage;
    oldWorker.onmessage = (event) => {
      if (event.data.id === writeId) (window as any).heldOldWrite = event;
      else handler.call(oldWorker, event);
    };
    (window as any).releaseOldWrite = () => handler.call(oldWorker, (window as any).heldOldWrite);
    const NativeWorker = window.Worker;
    class DelayedOpenWorker extends NativeWorker {
      private openId = -1;
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', (event) => {
          if (event.data.id === this.openId) {
            event.stopImmediatePropagation();
            (window as any).heldNewOpen = { worker: this, event };
          }
        });
      }
      postMessage(message: any, ...args: any[]) {
        if (message.method === 'open') this.openId = message.id;
        return Reflect.apply(super.postMessage, this, [message, ...args]);
      }
    }
    window.Worker = DelayedOpenWorker;
    (window as any).replacement = app
      .load(snapshot)
      .then(
        () => 'loaded',
        (error) => error.code,
      )
      .finally(() => {
        window.Worker = NativeWorker;
      });
    return old.id;
  });
  await page.waitForFunction(() => !!(window as any).heldNewOpen);
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    (window as any).oldWrite = book
      .getSheets()[0]
      .range('B2')
      .setValues([[71]]);
  });
  await page.waitForFunction(() => !!(window as any).heldOldWrite);
  await page.evaluate(() => {
    const held = (window as any).heldNewOpen;
    held.worker.onmessage.call(held.worker, held.event);
  });
  expect(await page.evaluate(() => (window as any).replacement)).toBe('BUSY');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(initialId);
  await expect(page.locator('#save-state')).toHaveText('Saving…');
  await page.evaluate(async () => {
    (window as any).releaseOldWrite();
    await (window as any).oldWrite;
  });
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual([[71]]);
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
});

test('a completed edit during replacement prevents overwriting the newer revision', async ({
  page,
}) => {
  const id = await page.evaluate(async () => {
    const app = window.opensheet;
    const old = app.getWorkbook();
    const snapshot = await old.toJSON();
    const prototype = Object.getPrototypeOf(old);
    const ready = prototype.ready;
    let intercepted = false;
    prototype.ready = function () {
      const result = ready.call(this);
      if (this === old || intercepted) return result;
      intercepted = true;
      (window as any).stagedId = this.id;
      return result.then(
        () =>
          new Promise((resolve) => {
            (window as any).releaseStaged = () => resolve(this);
          }),
      );
    };
    (window as any).replacement = app.load(snapshot).then(
      () => 'loaded',
      (error) => error.code,
    );
    return old.id;
  });
  await page.waitForFunction(() => !!(window as any).releaseStaged);
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('B2')
      .setValues([[72]]);
    (window as any).releaseStaged();
  });
  expect(await page.evaluate(() => (window as any).replacement)).toBe('BUSY');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual([[72]]);
  const exists = await page.evaluate(async () => {
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open('opensheet');
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const found = await new Promise<boolean>((resolve, reject) => {
      const request = db
        .transaction('heads')
        .objectStore('heads')
        .get((window as any).stagedId);
      request.onsuccess = () => resolve(request.result !== undefined);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return found;
  });
  expect(exists).toBe(false);
});

test('a completed edit prevents a delayed external workbook attachment', async ({ page }) => {
  const id = await page.evaluate(async () => {
    const current = window.opensheet.getWorkbook();
    const snapshot = await current.toJSON();
    const Workbook = current.constructor as new (input: typeof snapshot) => typeof current;
    const external = new Workbook({ ...snapshot, workbookId: `external_${crypto.randomUUID()}` });
    await external.ready();
    const ready = external.ready.bind(external);
    external.ready = async () => {
      await new Promise<void>((resolve) => {
        (window as any).releaseExternal = resolve;
      });
      return ready();
    };
    (window as any).external = external;
    (window as any).attachment = window.opensheet.attachWorkbook(external).then(
      () => 'attached',
      (error) => error.code,
    );
    return current.id;
  });
  await page.waitForFunction(() => !!(window as any).releaseExternal);
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('B2')
      .setValues([[73]]);
    (window as any).releaseExternal();
  });
  expect(await page.evaluate(() => (window as any).attachment)).toBe('BUSY');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual([[73]]);
  await page.evaluate(async () => {
    const external = (window as any).external;
    await external.getSheets()[0].range('A1').getValues();
    await external.deleteStorage();
    external.dispose();
  });
});

test('a completed edit prevents a delayed workbook restore', async ({ page }) => {
  const id = await page.evaluate(async () => {
    const app = window.opensheet;
    const current = app.getWorkbook();
    const snapshot = await current.toJSON();
    const Workbook = current.constructor as new (input: typeof snapshot) => typeof current;
    const target = new Workbook({ ...snapshot, workbookId: `restore_${crypto.randomUUID()}` });
    await target.ready();
    const targetId = target.id;
    await target.close();
    const NativeWorker = window.Worker;
    class DelayedOpenWorker extends NativeWorker {
      private openId = -1;
      constructor(url: string | URL, options?: WorkerOptions) {
        super(url, options);
        this.addEventListener('message', (event) => {
          if (event.data.id === this.openId) {
            event.stopImmediatePropagation();
            (window as any).heldRestore = { worker: this, event };
          }
        });
      }
      postMessage(message: any, ...args: any[]) {
        if (message.method === 'open') this.openId = message.id;
        return Reflect.apply(super.postMessage, this, [message, ...args]);
      }
    }
    window.Worker = DelayedOpenWorker;
    (window as any).restoration = app
      .open(targetId)
      .then(
        () => 'opened',
        (error) => error.code,
      )
      .finally(() => {
        window.Worker = NativeWorker;
      });
    return current.id;
  });
  await page.waitForFunction(() => !!(window as any).heldRestore);
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('B2')
      .setValues([[74]]);
    const held = (window as any).heldRestore;
    held.worker.onmessage.call(held.worker, held.event);
  });
  expect(await page.evaluate(() => (window as any).restoration)).toBe('BUSY');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().id)).toBe(id);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual([[74]]);
});

test('a rejected write remains visible through later successes until acknowledged', async ({
  page,
}) => {
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    let fail = true;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute' && fail) {
        fail = false;
        queueMicrotask(() =>
          handler.call(
            worker,
            new MessageEvent('message', {
              data: {
                id: message.id,
                error: { code: 'STORAGE_BUDGET', message: 'Application storage budget exceeded' },
              },
            }),
          ),
        );
        return;
      }
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    const sheet = book.getSheets()[0];
    const previous = await sheet.range('B2').getValues();
    const revision = book.revision;
    try {
      await sheet.range('B2').setValues([[123]]);
    } catch {}
    return { previous, revision };
  });
  await expect(page.locator('#save-state')).toHaveText('Save unconfirmed · Storage full');
  await expect(page.getByRole('button', { name: 'Dismiss', exact: true })).toBeVisible();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before.revision);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('B2').getValues(),
    ),
  ).toEqual(before.previous);
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('B3')
      .setValues([[77]]),
  );
  await expect(page.locator('#save-state')).toHaveText('Save unconfirmed · Storage full');
  await page.getByRole('button', { name: 'Dismiss', exact: true }).click();
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
});

test('read-only state and unrecoverable worker failures are explicit', async ({ page }) => {
  const second = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const Constructor = book.constructor as new (
      snapshot: undefined,
      options: object,
      id: string,
    ) => typeof book;
    const reader = await new Constructor(undefined, {}, book.id).ready();
    try {
      return reader.persistenceState;
    } finally {
      await reader.close();
    }
  });
  expect(second.state).toBe('read-only');
  await page.evaluate(() => window.opensheet.setMode('read'));
  await expect(page.locator('#save-state')).toHaveText('Read-only');
  await page.evaluate(() => window.opensheet.setMode('edit'));
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
  await page.evaluate(() => {
    const worker = (window.opensheet.getWorkbook() as any).worker as Worker;
    worker.onerror!(new ErrorEvent('error', { message: 'worker interrupted' }));
  });
  await expect(page.locator('#save-state')).toHaveText('Save unconfirmed · Reopen workbook');
  await expect(page.getByRole('button', { name: 'Dismiss', exact: true })).toBeHidden();
});

test('a fatal Worker failure preserves the current cell draft when commit is attempted', async ({
  page,
}) => {
  const initial = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['persisted']]);
    window.opensheet.selection.set({
      sheetId: sheet.id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
    return book.id;
  });
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  await page.keyboard.type('unsaved draft');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await expect(editor).toHaveValue('unsaved draft');
  await page.evaluate(() => {
    const worker = (window.opensheet.getWorkbook() as any).worker as Worker;
    worker.onerror?.call(worker, new ErrorEvent('error', { message: 'Worker interrupted' }));
  });
  await expect(page.locator('#save-state')).toHaveText('Save unconfirmed · Reopen workbook');
  await editor.press('Enter');
  await expect(editor).toBeVisible();
  await expect(editor).toHaveValue('unsaved draft');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  await editor.press('Escape');
  await page.evaluate(async (id) => window.opensheet.open(id), initial);
  expect(
    await page.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['persisted']]);
});

test('uncommitted formula drafts are visible and prevent replacement until cancelled', async ({
  page,
}) => {
  const snapshot = await page.evaluate(async () =>
    JSON.stringify(await window.opensheet.getWorkbook().toJSON()),
  );
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  const formula = page.locator('.os-formula-input');
  await formula.fill('=1+2');
  await expect(page.locator('#save-state')).toHaveText('Editing…');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
  expect(
    await page.evaluate(() => {
      const event = new Event('beforeunload', { cancelable: true });
      window.dispatchEvent(event);
      return event.defaultPrevented;
    }),
  ).toBe(true);
  expect(
    await page.evaluate(async (file) => {
      try {
        await window.opensheet.load(JSON.parse(file));
        return 'loaded';
      } catch (error) {
        return (error as { code?: string }).code;
      }
    }, snapshot),
  ).toBe('BUSY');
  await formula.press('Escape');
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
});

test('cell drafts stay distinct from persisted data and clear on cancellation', async ({
  page,
}) => {
  const grid = page.getByRole('grid', { name: 'Spreadsheet' });
  await grid.focus();
  await page.keyboard.press('F2');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await editor.fill('draft text');
  await expect(page.locator('#save-state')).toHaveText('Editing…');
  await editor.press('Escape');
  await expect(page.locator('#save-state')).toHaveText('Saved locally');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(false);
});

test('import prevents new drafts until the replacement is attached', async ({ page }) => {
  await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const snapshot = await book.toJSON();
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    let importId: number | undefined;
    (window as any).importHeld = false;
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'importJSON') importId = message.id;
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (event.data.id === importId) {
        (window as any).importHeld = true;
        (window as any).releaseImport = () => handler.call(worker, event);
      } else handler.call(worker, event);
    };
    (window as any).importPromise = app.loadJSON(snapshot);
  });
  await page.waitForFunction(() => (window as any).importHeld);
  const formula = page.locator('.os-formula-input');
  await expect(formula).toHaveAttribute('readonly', '');
  const grid = page.getByRole('grid', { name: 'Spreadsheet' });
  await expect(grid).toHaveAttribute('aria-readonly', 'true');
  await grid.focus();
  await page.keyboard.type('x');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(false);
  await page.evaluate(async () => {
    (window as any).releaseImport();
    await (window as any).importPromise;
  });
  await expect(formula).not.toHaveAttribute('readonly', '');
  await expect(grid).toHaveAttribute('aria-readonly', 'false');
  await formula.fill('=1+2');
  await expect(page.locator('#save-state')).toHaveText('Editing…');
});

test('failed imports restore editability without changing the workbook', async ({ page }) => {
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  const result = await page.evaluate(async () => {
    try {
      await window.opensheet.loadJSON('{broken');
      return 'imported';
    } catch {
      return 'rejected';
    }
  });
  expect(result).toBe('rejected');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  const formula = page.locator('.os-formula-input');
  await expect(formula).not.toHaveAttribute('readonly', '');
  await formula.fill('=2+3');
  await expect(page.locator('#save-state')).toContainText('Save unconfirmed');
  expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
});
