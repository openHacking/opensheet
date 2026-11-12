import { expect, test } from '@playwright/test';
import type { OpenSheet } from 'opensheet';

test('transaction scopes cannot close or delete the parent workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[10]]);
    const codes: string[] = [];
    let expired: typeof book | undefined;
    await book.transaction({}, async (scope) => {
      expired = scope;
      for (const operation of [
        () => scope.deleteStorage(),
        () => scope.close(),
        () => book.deleteStorage(),
        () => book.close(),
        () => scope.generate(100),
        () => book.generate(100),
      ]) {
        try {
          await operation();
          codes.push('unexpected success');
        } catch (error) {
          codes.push((error as any).code);
        }
      }
      await scope
        .getSheets()[0]
        .range('A1')
        .setValues([[20]]);
    });
    for (const operation of [() => expired!.deleteStorage(), () => expired!.close()]) {
      try {
        await operation();
        codes.push('unexpected success');
      } catch (error) {
        codes.push((error as any).code);
      }
    }
    const value = await sheet.range('A1').getValues();
    const revision = book.revision;
    await app.open(book.id);
    const reopened = app.getWorkbook();
    const persisted = await reopened.getSheets()[0].range('A1').getValues();
    const undo = await reopened.undo();
    const undone = await reopened.getSheets()[0].range('A1').getValues();
    const redo = await reopened.redo();
    const redone = await reopened.getSheets()[0].range('A1').getValues();
    return { codes, value, revision, persisted, undo, undone, redo, redone };
  });
  expect(result).toEqual({
    codes: ['BUSY', 'BUSY', 'BUSY', 'BUSY', 'BUSY', 'BUSY', 'DISPOSED', 'DISPOSED'],
    value: [[20]],
    revision: result.revision,
    persisted: [[20]],
    undo: true,
    undone: [[10]],
    redo: true,
    redone: [[20]],
  });
});

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
});

test('keeps data subscriptions bounded through 100 same-workbook reattachments', async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const host = document.createElement('div');
    host.style.height = '300px';
    document.body.append(host);
    const Editor = window.opensheet.constructor as new (options: {
      container: HTMLElement;
    }) => OpenSheet;
    const editor = new Editor({ container: host });
    try {
      const book = await editor.createWorkbook();
      const subscribe = book.onData.bind(book);
      let active = 0;
      book.onData = (listener) => {
        active++;
        const stop = subscribe(listener);
        let stopped = false;
        return () => {
          if (!stopped) {
            stopped = true;
            active--;
            stop();
          }
        };
      };
      const counts = [];
      for (let i = 0; i < 100; i++) {
        await editor.attachWorkbook(book);
        counts.push(active);
      }
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([['still usable']]);
      const values = await book.getSheets()[0].range('A1').getValues();
      editor.dispose();
      return { counts, active, values, grids: host.querySelectorAll('.os-grid').length };
    } finally {
      editor.dispose();
      host.remove();
    }
  });
  expect(result.counts).toEqual(Array(100).fill(2)); // Grid and editor each own one data subscription.
  expect(result.active).toBe(0);
  expect(result.values).toEqual([['still usable']]);
  expect(result.grids).toBe(0);
});

test('does not attach a pending external workbook after the editor is disposed', async ({
  page,
}) => {
  const result = await page.evaluate(async () => {
    const host = document.createElement('div');
    host.style.height = '300px';
    document.body.append(host);
    const Editor = window.opensheet.constructor as new (options: {
      container: HTMLElement;
    }) => OpenSheet;
    const editor = new Editor({ container: host });
    const external = window.opensheet.getWorkbook();
    const ready = external.ready.bind(external);
    let release!: () => void;
    const pending = new Promise<void>((resolve) => {
      release = resolve;
    });
    external.ready = async () => {
      await pending;
      return ready();
    };
    try {
      const attachment = editor.attachWorkbook(external).then(
        () => 'attached',
        (error) => error.code,
      );
      editor.dispose();
      release();
      const status = await attachment;
      return {
        status,
        grids: editor.element.querySelectorAll('.os-grid').length,
        values: await external.getSheets()[0].range('F11').getValues(),
      };
    } finally {
      external.ready = ready;
      release();
      host.remove();
    }
  });
  expect(result.status).toBe('DISPOSED');
  expect(result.grids).toBe(0);
  expect(result.values).toEqual([[13097]]);
});

for (const kind of ['json', 'binary'] as const) {
  test(`a late ${kind} import cannot replace a newer workbook without a cancellation signal`, async ({
    page,
  }) => {
    const result = await page.evaluate(async (kind) => {
      const app = window.opensheet;
      const old = app.getWorkbook();
      const snapshot = await old.toJSON();
      const Workbook = old.constructor as new (input: typeof snapshot) => typeof old;
      const next = new Workbook({ ...snapshot, workbookId: `next_${crypto.randomUUID()}` });
      await next.ready();
      let release!: () => void;
      const pending = new Promise<void>((resolve) => {
        release = resolve;
      });
      if (kind === 'json') old.importJSON = async () => pending;
      else old.importBinary = async () => pending;
      const importing = (
        kind === 'json' ? app.loadJSON('{}') : app.loadBinary(new Uint8Array())
      ).then(
        () => 'loaded',
        (error) => error.code,
      );
      await app.attachWorkbook(next);
      await next
        .getSheets()[0]
        .range('A1')
        .setValues([['new workbook']]);
      release();
      const status = await importing;
      const current = app.getWorkbook();
      return {
        status,
        id: current.id,
        expectedId: next.id,
        value: await current.getSheets()[0].range('A1').getValues(),
      };
    }, kind);
    expect(result).toEqual({
      status: 'BUSY',
      id: result.expectedId,
      expectedId: result.expectedId,
      value: [['new workbook']],
    });
  });
}

for (const kind of ['create', 'load'] as const) {
  test(`cleans up a ${kind} that finishes after its editor is disposed`, async ({ page }) => {
    const result = await page.evaluate(async (kind) => {
      const host = document.createElement('div');
      document.body.append(host);
      const Editor = window.opensheet.constructor as new (options: {
        container: HTMLElement;
      }) => OpenSheet;
      const editor = new Editor({ container: host });
      const external = window.opensheet.getWorkbook();
      const snapshot = kind === 'load' ? await external.toJSON() : undefined;
      const prototype = Object.getPrototypeOf(external);
      const originalReady = prototype.ready;
      let release!: () => void;
      let initialized!: (id: string) => void;
      const blocked = new Promise<void>((resolve) => {
        release = resolve;
      });
      const started = new Promise<string>((resolve) => {
        initialized = resolve;
      });
      prototype.ready = async function () {
        await originalReady.call(this);
        initialized(this.id);
        await blocked;
        return this;
      };
      try {
        const operation = (
          kind === 'create' ? editor.createWorkbook() : editor.load(snapshot!)
        ).then(
          () => 'attached',
          (error: any) => error.code,
        );
        const id = await started;
        editor.dispose();
        release();
        const status = await operation;
        const request = indexedDB.open('opensheet');
        const db = await new Promise<IDBDatabase>((resolve, reject) => {
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const range = IDBKeyRange.bound(`${id}/`, `${id}/\uffff`);
        const requests = [
          tx.objectStore('heads').get(id),
          ...['records', 'sizes', 'scratch'].map((name) => tx.objectStore(name).count(range)),
        ];
        const values = await Promise.all(
          requests.map(
            (item) =>
              new Promise<any>((resolve, reject) => {
                item.onsuccess = () => resolve(item.result);
                item.onerror = () => reject(item.error);
              }),
          ),
        );
        db.close();
        return {
          status,
          hasHead: values[0] !== undefined,
          counts: values.slice(1),
          external: await external.getSheets()[0].range('F11').getValues(),
        };
      } finally {
        prototype.ready = originalReady;
        release();
        editor.dispose();
        host.remove();
      }
    }, kind);
    expect(result).toEqual({
      status: 'DISPOSED',
      hasHead: false,
      counts: [0, 0, 0],
      external: [[13097]],
    });
  });
}
