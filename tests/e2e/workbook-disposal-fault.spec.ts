import { expect, test } from '@playwright/test';

for (const failure of ['channel', 'worker'] as const) {
  test(`workbook disposal releases pending requests after ${failure} cleanup fails`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (failure) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[421]]);
      const internal = book as any,
        worker = internal.worker as Worker,
        channel = internal.channel as BroadcastChannel;
      const handler = worker.onmessage!,
        post = worker.postMessage.bind(worker),
        terminate = worker.terminate.bind(worker),
        close = channel.close.bind(channel);
      let statsId = -1,
        held = false,
        terminated = false,
        notifications = 0;
      book.onDispose(() => {
        notifications++;
      });
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'stats') statsId = message.id;
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === statsId) held = true;
        else handler.call(worker, event);
      };
      worker.terminate = () => {
        terminated = true;
        terminate();
        if (failure === 'worker') throw new Error('Injected worker termination failure');
      };
      channel.close = () => {
        close();
        if (failure === 'channel') throw new Error('Injected channel close failure');
      };
      const pending = book.storageStats().then(
        () => 'resolved',
        (error) => error.code,
      );
      for (let i = 0; i < 100 && !held; i++)
        await new Promise((resolve) => setTimeout(resolve, 10));
      if (!held) throw new Error('Expected stats response was not held');
      let error = '';
      try {
        book.dispose();
      } catch (reason) {
        error = String(reason);
      }
      const pendingResult = await Promise.race([
        pending,
        new Promise<string>((resolve) => setTimeout(() => resolve('timeout'), 2000)),
      ]);
      let second = '';
      try {
        book.dispose();
      } catch (reason) {
        second = String(reason);
      }
      return {
        id: book.id,
        error,
        second,
        pendingResult,
        terminated,
        notifications,
        pendingCount: internal.pending.size,
      };
    }, failure);
    expect(result).toEqual({
      id: expect.any(String),
      error: `Error: Injected ${failure === 'channel' ? 'channel close' : 'worker termination'} failure`,
      second: '',
      pendingResult: 'DISPOSED',
      terminated: true,
      notifications: 1,
      pendingCount: 0,
    });
    const reopened = await context.newPage();
    await reopened.goto('/#budget');
    await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
    expect(
      await reopened.evaluate(async () => ({
        id: window.opensheet.getWorkbook().id,
        readOnly: window.opensheet.getWorkbook().isReadOnly,
        value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      })),
    ).toEqual({ id: result.id, readOnly: false, value: [[421]] });
    await reopened.close();
  });
}

test('a cleanup failure does not misreport an already committed workbook deletion', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      channel = (book as any).channel as BroadcastChannel;
    const close = channel.close.bind(channel),
      log = console.error;
    const logged: string[] = [];
    channel.close = () => {
      close();
      throw new Error('Injected channel close failure');
    };
    console.error = (...args) => {
      logged.push(args.map(String).join(' '));
      throw new Error('Injected logging failure');
    };
    let outcome = '';
    try {
      await book.deleteStorage();
      outcome = 'deleted';
    } catch (error) {
      outcome = String(error);
    } finally {
      console.error = log;
    }
    const request = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    const read = db.transaction('heads').objectStore('heads').get(book.id);
    const head = await new Promise<unknown>((resolve, reject) => {
      read.onsuccess = () => resolve(read.result);
      read.onerror = () => reject(read.error);
    });
    db.close();
    return { outcome, headExists: head !== undefined, logged };
  });
  expect(result.outcome).toBe('deleted');
  expect(result.headExists).toBe(false);
  expect(result.logged.some((entry) => entry.includes('Injected channel close failure'))).toBe(
    true,
  );
});

test('close preserves its original RPC error when cleanup also fails', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      internal = book as any,
      channel = internal.channel as BroadcastChannel;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[532]]);
    const rpc = internal.rpc.bind(book),
      close = channel.close.bind(channel),
      log = console.error;
    const logged: string[] = [];
    internal.rpc = (method: string, ...args: unknown[]) =>
      method === 'close'
        ? Promise.reject(new Error('Injected close RPC failure'))
        : rpc(method, ...args);
    channel.close = () => {
      close();
      throw new Error('Injected channel close failure');
    };
    console.error = (...args) => {
      logged.push(args.map(String).join(' '));
    };
    let outcome = '';
    try {
      await book.close();
    } catch (error) {
      outcome = String(error);
    } finally {
      console.error = log;
    }
    return { id: book.id, outcome, logged };
  });
  expect(result.outcome).toBe('Error: Injected close RPC failure');
  expect(result.logged.some((entry) => entry.includes('Injected channel close failure'))).toBe(
    true,
  );
  const reopened = await context.newPage();
  await reopened.goto('/#budget');
  await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await reopened.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      readOnly: window.opensheet.getWorkbook().isReadOnly,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    })),
  ).toEqual({ id: result.id, readOnly: false, value: [[532]] });
});

test('multiple workbook cleanup failures preserve the first error when logging fails', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      internal = book as any;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[642]]);
    const channel = internal.channel as BroadcastChannel,
      worker = internal.worker as Worker;
    const close = channel.close.bind(channel),
      terminate = worker.terminate.bind(worker),
      log = console.error;
    channel.close = () => {
      close();
      throw undefined;
    };
    worker.terminate = () => {
      terminate();
      throw new Error('Second cleanup failure');
    };
    let caught = false,
      reason: unknown = 'not thrown',
      logs = 0;
    console.error = () => {
      logs++;
      throw new Error('Injected logging failure');
    };
    try {
      book.dispose();
    } catch (error) {
      caught = true;
      reason = error;
    } finally {
      console.error = log;
    }
    return {
      id: book.id,
      caught,
      reasonWasUndefined: reason === undefined,
      logs,
      pending: internal.pending.size,
    };
  });
  expect(result).toEqual({
    id: expect.any(String),
    caught: true,
    reasonWasUndefined: true,
    logs: 1,
    pending: 0,
  });
  const reopened = await context.newPage();
  await reopened.goto('/#budget');
  await expect(reopened.locator('#scene-title')).toHaveText('Budget Simulator');
  expect(
    await reopened.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      readOnly: window.opensheet.getWorkbook().isReadOnly,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    })),
  ).toEqual({ id: result.id, readOnly: false, value: [[642]] });
});
