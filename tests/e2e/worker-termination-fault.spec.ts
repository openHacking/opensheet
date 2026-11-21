import { expect, test } from '@playwright/test';

for (const kind of ['error', 'messageerror', 'response'] as const) {
  test(`worker ${kind} rejects pending requests even if termination throws`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[713]]);
      const internal = book as any,
        worker = internal.worker as Worker;
      const terminate = worker.terminate.bind(worker),
        post = worker.postMessage.bind(worker),
        handler = worker.onmessage!;
      let id = -1,
        logged = 0;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'stats') id = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = () => {};
      worker.terminate = () => {
        terminate();
        throw new Error('termination failed');
      };
      const originalLog = console.error;
      console.error = () => {
        logged++;
        throw new Error('logger failed');
      };
      try {
        const pending = book.storageStats().then(
          () => 'resolved',
          (error) => error.code,
        );
        if (kind === 'error')
          worker.onerror?.call(worker, new ErrorEvent('error', { message: 'Injected failure' }));
        else if (kind === 'messageerror')
          worker.onmessageerror?.call(worker, new MessageEvent('messageerror'));
        else {
          // Invoke the real dispatcher with a failed response, while another request remains pending.
          internal.pending.set(id + 1000, { resolve() {}, reject() {}, timer: 0 });
          handler.call(
            worker,
            new MessageEvent('message', {
              data: {
                id: id + 1000,
                error: { code: 'WORKER_FAILED', message: 'Injected failure' },
              },
            }),
          );
        }
        const code = await Promise.race([
          pending,
          new Promise<string>((resolve) => setTimeout(() => resolve('timeout'), 1000)),
        ]);
        return { code, logged, pending: internal.pending.size };
      } finally {
        console.error = originalLog;
      }
    }, kind);
    expect(result).toEqual({ code: 'WORKER_FAILED', logged: 1, pending: 0 });
    const reopened = await context.newPage();
    await reopened.goto('/#budget');
    await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
    expect(
      await reopened.evaluate(async () => ({
        readOnly: window.opensheet.getWorkbook().isReadOnly,
        value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      })),
    ).toEqual({ readOnly: false, value: [[713]] });
    await reopened.close();
  });
}

test('initialization preserves the worker error when termination and disposal fail', async ({
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
      .setValues([[714]]);
    const id = original.id;
    await original.close();
    const WorkerConstructor = globalThis.Worker,
      originalLog = console.error;
    let logged = 0;
    try {
      globalThis.Worker = class extends WorkerConstructor {
        override postMessage(message: any) {
          if (message.method === 'open')
            queueMicrotask(() =>
              this.onerror?.call(this, new ErrorEvent('error', { message: 'startup failure' })),
            );
        }
        override terminate() {
          super.terminate();
          throw new Error('termination failed');
        }
      };
      console.error = () => {
        logged++;
        throw new Error('logger failed');
      };
      let failure = { code: '', message: '' };
      try {
        await openWorkbook(id);
      } catch (error) {
        failure = { code: (error as any).code, message: (error as Error).message };
      }
      return { failure, logged };
    } finally {
      globalThis.Worker = WorkerConstructor;
      console.error = originalLog;
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.failure).toEqual({
    code: 'WORKER_FAILED',
    message: expect.stringContaining('startup failure'),
  });
  expect(result.logged).toBe(2);
});

test('committed edit remains acknowledged if notification and termination fail', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      internal = book as any;
    const worker = internal.worker as Worker,
      terminate = worker.terminate.bind(worker),
      originalLog = console.error;
    internal.channel.postMessage = () => {
      throw new Error('notification failed');
    };
    worker.terminate = () => {
      terminate();
      throw new Error('termination failed');
    };
    let logged = 0;
    console.error = () => {
      logged++;
      throw new Error('logger failed');
    };
    try {
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[715]]);
      const next = await book.storageStats().then(
        () => 'resolved',
        (error) => error.code,
      );
      return { next, logged };
    } finally {
      console.error = originalLog;
    }
  });
  expect(result).toEqual({ next: 'SYNC_UNAVAILABLE', logged: 1 });
  const reopened = await context.newPage();
  await reopened.goto('/#budget');
  await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await reopened.evaluate(async () => ({
      readOnly: window.opensheet.getWorkbook().isReadOnly,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    })),
  ).toEqual({ readOnly: false, value: [[715]] });
  await reopened.close();
});

test('a timed-out request still fails its peers when worker termination throws', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      internal = book as any;
    const worker = internal.worker as Worker,
      terminate = worker.terminate.bind(worker),
      originalLog = console.error;
    internal.options.operationTimeoutMs = 100;
    worker.onmessage = () => {};
    worker.terminate = () => {
      terminate();
      throw new Error('termination failed');
    };
    console.error = () => {
      throw new Error('logger failed');
    };
    try {
      const a = book.storageStats().then(
        () => 'resolved',
        (error) => error.code,
      );
      const b = book.storageStats().then(
        () => 'resolved',
        (error) => error.code,
      );
      const codes = await Promise.race([
        Promise.all([a, b]),
        new Promise<string[]>((resolve) => setTimeout(() => resolve(['timeout']), 2000)),
      ]);
      return { codes, pending: internal.pending.size };
    } finally {
      console.error = originalLog;
    }
  });
  expect(result).toEqual({ codes: ['WORKER_TIMEOUT', 'WORKER_FAILED'], pending: 0 });
});
