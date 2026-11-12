import { expect, test } from '@playwright/test';

test('notification failure leaves another committed but unacknowledged write unconfirmed', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    return {
      revision: book.revision,
      a65: (await book.getSheets()[0].range('A65').getValues())[0][0],
    };
  });
  const reader = await context.newPage();
  await reader.goto('/#budget');
  await expect(reader.locator('#take-write-lock')).toBeVisible();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const worker = (book as any).worker as Worker;
    const channel = (book as any).channel as BroadcastChannel;
    const originalPost = worker.postMessage.bind(worker),
      originalHandler = worker.onmessage!;
    const ids: number[] = [];
    const responses = new Map<number, MessageEvent>();
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') ids.push(message.id);
      return Reflect.apply(originalPost, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (!ids.includes(event.data.id)) return originalHandler.call(worker, event);
      responses.set(event.data.id, event);
      if (responses.size === 2) {
        channel.postMessage = () => {
          throw new DOMException('Closed', 'InvalidStateError');
        };
        originalHandler.call(worker, responses.get(ids[0])!);
      }
    };
    const first = book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 111 } }],
      },
    });
    const second = book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 64, column: 0, input: { type: 'number', value: 222 } }],
      },
    });
    const [a, b] = await Promise.allSettled([first, second]);
    return {
      responseCount: responses.size,
      bothCommitted: [...responses.values()].every(
        (event) => !event.data.error && !!event.data.result?.commit,
      ),
      first: a.status,
      second:
        b.status === 'rejected' ? { code: b.reason.code, message: b.reason.message } : b.status,
      revision: book.revision,
      state: book.persistenceState,
    };
  });
  expect(result.responseCount).toBe(2);
  expect(result.bothCommitted).toBe(true);
  expect(result.first).toBe('fulfilled');
  expect(result.second).toEqual({
    code: 'SYNC_UNAVAILABLE',
    message: 'Tab synchronization failed. Reopen the workbook and check pending changes.',
  });
  expect(result.revision).toBe(initial.revision + 1);
  expect(result.state.state).toBe('failed');
  await expect(page.locator('#save-state')).toHaveText('Sync unavailable · Reopen workbook');
  await expect
    .poll(() => reader.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()), {
      timeout: 10000,
    })
    .toBe(true);
  const recovered = await reader.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    const read = async () => [
      (await sheet.range('A1').getValues())[0][0],
      (await sheet.range('A65').getValues())[0][0],
    ];
    const revision = book.revision,
      values = await read();
    const undo = await book.undo(),
      undone = await read();
    const redo = await book.redo(),
      redone = await read();
    return { revision, values, undo, undone, redo, redone };
  });
  expect(recovered).toEqual({
    revision: initial.revision + 2,
    values: [111, 222],
    undo: true,
    undone: [111, initial.a65],
    redo: true,
    redone: [111, 222],
  });
  await reader.close();
});

for (const operation of ['edit', 'undo', 'redo', 'generate', 'json', 'binary'] as const) {
  test(`committed ${operation} succeeds when tab notification fails`, async ({ page, context }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const before = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[42]]);
      if (kind === 'redo') await book.undo();
      let input: unknown;
      if (kind === 'json') input = await book.toJSON();
      if (kind === 'binary') {
        const chunks: BlobPart[] = [];
        for await (const chunk of book.streamBinary()) chunks.push(chunk as BlobPart);
        input = new Blob(chunks);
      }
      (window as any).syncFaultInput = input;
      return { id: book.id, revision: book.revision };
    }, operation);
    const reader = await context.newPage();
    await reader.goto('/#budget');
    await expect(reader.locator('#take-write-lock')).toBeVisible();
    const result = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook();
      const channel = (book as any).channel as BroadcastChannel;
      channel.postMessage = () => {
        throw new DOMException('Closed', 'InvalidStateError');
      };
      let value: unknown,
        error = '';
      try {
        if (kind === 'edit')
          value = await book
            .getSheets()[0]
            .range('A1')
            .setValues([[77]]);
        else if (kind === 'undo') value = await book.undo();
        else if (kind === 'redo') value = await book.redo();
        else if (kind === 'generate') value = await book.generate(100);
        else if (kind === 'json') value = await book.importJSON((window as any).syncFaultInput);
        else value = await book.importBinary((window as any).syncFaultInput);
      } catch (reason) {
        error = (reason as Error).message;
      }
      let nextCode = '';
      try {
        await book.storageStats();
      } catch (reason) {
        nextCode = (reason as { code?: string }).code ?? '';
      }
      return {
        error,
        successful: kind === 'undo' || kind === 'redo' ? value === true : true,
        revision: book.revision,
        state: book.persistenceState.state,
        code: book.persistenceState.error?.code,
        nextCode,
      };
    }, operation);
    expect(result).toEqual({
      error: '',
      successful: true,
      revision: before.revision + 1,
      state: 'failed',
      code: 'SYNC_UNAVAILABLE',
      nextCode: 'SYNC_UNAVAILABLE',
    });
    await expect(page.locator('#save-state')).toHaveText('Sync unavailable · Reopen workbook');
    await expect
      .poll(() => reader.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()), {
        timeout: 10000,
      })
      .toBe(true);
    const reopened = await reader.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        revision: book.revision,
        value: (await book.getSheets()[0].range('A1').getValues())[0][0],
      };
    });
    expect(reopened.revision).toBe(before.revision + 1);
    if (operation === 'edit') expect(reopened.value).toBe(77);
    if (operation === 'redo' || operation === 'json' || operation === 'binary')
      expect(reopened.value).toBe(42);
    await reader.close();
  });
}
