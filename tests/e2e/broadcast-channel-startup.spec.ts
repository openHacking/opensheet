import { expect, test } from '@playwright/test';

test('denied tab synchronization releases the writer and allows the saved workbook to reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (modulePath) => {
    const { openWorkbook } = await import(/* @vite-ignore */ modulePath);
    const original = window.opensheet.getWorkbook();
    await original
      .getSheets()[0]
      .range('A1')
      .setValues([[12345]]);
    const id = original.id;
    const revision = original.revision;
    await original.close();

    const Channel = globalThis.BroadcastChannel;
    const failures: { code?: string; name: string; message: string }[] = [];
    try {
      for (const error of [
        new DOMException('Blocked', 'SecurityError'),
        new DOMException('Blocked', 'NotAllowedError'),
        new TypeError('Unexpected constructor defect'),
      ]) {
        globalThis.BroadcastChannel = class extends Channel {
          constructor(name: string) {
            super(name);
            this.close();
            throw error;
          }
        };
        try {
          await openWorkbook(id);
          throw new Error('Expected channel construction to fail');
        } catch (failure) {
          const observed = failure as Error & { code?: string };
          failures.push({ code: observed.code, name: observed.name, message: observed.message });
        }
      }
    } finally {
      globalThis.BroadcastChannel = Channel;
    }
    const reopened = await openWorkbook(id);
    const recovered = {
      revision: reopened.revision,
      writer: !reopened.isReadOnly,
      value: await reopened.getSheets()[0].range('A1').getValues(),
    };
    await reopened.close();
    return { failures, revision, recovered };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);

  expect(result.failures[0]).toMatchObject({
    code: 'BROADCAST_CHANNEL_UNAVAILABLE',
    message: expect.stringContaining('reopen the workbook'),
  });
  expect(result.failures[1]).toMatchObject({ code: 'BROADCAST_CHANNEL_UNAVAILABLE' });
  expect(result.failures[2]).toMatchObject({
    name: 'TypeError',
    message: 'Unexpected constructor defect',
  });
  expect(result.recovered).toEqual({ revision: result.revision, writer: true, value: [[12345]] });
});

test('channel denial leaves no new workbook, and Worker denial closes the channel', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (modulePath) => {
    const { createWorkbookFile, Workbook } = await import(/* @vite-ignore */ modulePath);
    const file = createWorkbookFile({ sheets: [{ name: 'New', rows: 2, columns: 2 }] });
    const req = <T>(request: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
    const db = await req(indexedDB.open('opensheet'));
    const keys = async () => {
      const tx = db.transaction(['heads', 'records', 'sizes', 'scratch'], 'readonly');
      return Promise.all(
        ['heads', 'records', 'sizes', 'scratch'].map((name) =>
          req(tx.objectStore(name).getAllKeys()),
        ),
      );
    };
    const before = await keys();
    const OriginalWorker = globalThis.Worker;
    const OriginalChannel = globalThis.BroadcastChannel;
    let workerStarts = 0;
    let closed = 0;
    let channelError: string | undefined;
    let workerError: string | undefined;
    try {
      globalThis.Worker = class extends OriginalWorker {
        constructor(...args: ConstructorParameters<typeof Worker>) {
          workerStarts++;
          super(...args);
        }
      };
      globalThis.BroadcastChannel = class extends OriginalChannel {
        constructor(name: string) {
          super(name);
          super.close();
          throw new DOMException('Denied', 'SecurityError');
        }
      };
      try {
        await new Workbook(file).ready();
      } catch (error) {
        channelError = (error as { code?: string }).code;
      }
      const after = await keys();
      globalThis.BroadcastChannel = class extends OriginalChannel {
        close() {
          closed++;
          super.close();
        }
      };
      globalThis.Worker = class extends OriginalWorker {
        constructor(...args: ConstructorParameters<typeof Worker>) {
          super(...args);
          super.terminate();
          throw new DOMException('Denied', 'SecurityError');
        }
      };
      try {
        await new Workbook(file).ready();
      } catch (error) {
        workerError = (error as { code?: string }).code;
      }
      return { before, after, workerStarts, closed, channelError, workerError };
    } finally {
      globalThis.Worker = OriginalWorker;
      globalThis.BroadcastChannel = OriginalChannel;
      db.close();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.channelError).toBe('BROADCAST_CHANNEL_UNAVAILABLE');
  expect(result.workerError).toBe('WORKER_UNAVAILABLE');
  expect(result.workerStarts).toBe(0);
  expect(result.after).toEqual(result.before);
  expect(result.closed).toBe(1);
});
