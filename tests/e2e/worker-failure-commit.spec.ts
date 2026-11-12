import { expect, test } from '@playwright/test';

for (const failure of ['error', 'messageerror'] as const) {
  test(`Worker ${failure} during head transaction preserves the prior commit`, async ({
    page,
    context,
  }) => {
    await page.route('**/engine.worker.js*', async (route) => {
      const response = await route.fetch();
      await route.fulfill({
        response,
        body:
          (await response.text()) +
          `
let armHeadFailure = false;
addEventListener('message', event => { if (event.data.armHeadFailure) armHeadFailure = true; });
const originalPutForFailure = IDBObjectStore.prototype.put;
IDBObjectStore.prototype.put = function(...args) {
  const request = originalPutForFailure.apply(this, args);
  if (armHeadFailure && this.name === 'heads') {
    armHeadFailure = false;
    const tx = this.transaction;
    request.addEventListener('success', () => {
      const keepAlive = () => { const next = tx.objectStore('heads').get('__worker_failure_keepalive__'); next.onsuccess = keepAlive; };
      keepAlive();
      postMessage({ workerFailureStage: 'head-transaction' });
    });
  }
  return request;
};`,
      });
    });
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const initial = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      const original = [
        (await sheet.range('A1').getValues())[0][0],
        (await sheet.range('A65').getValues())[0][0],
      ];
      await book.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [
            { row: 0, column: 0, input: { type: 'number', value: 10 } },
            { row: 64, column: 0, input: { type: 'number', value: 20 } },
          ],
        },
      });
      return { revision: book.revision, original };
    });
    const reader = await context.newPage();
    await reader.goto('/#budget');
    await expect(reader.locator('#take-write-lock')).toBeVisible();
    const result = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook(),
        worker = (book as any).worker as Worker;
      let reached = false;
      worker.addEventListener('message', (event) => {
        if (event.data.workerFailureStage !== 'head-transaction') return;
        reached = true;
        if (kind === 'error')
          worker.onerror?.call(
            worker,
            new ErrorEvent('error', { message: 'Injected during commit' }),
          );
        else worker.onmessageerror?.call(worker, new MessageEvent('messageerror'));
      });
      worker.postMessage({ armHeadFailure: true });
      let code = '';
      try {
        await book.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: book.getSheets()[0].id,
            cells: [
              { row: 0, column: 0, input: { type: 'number', value: 101 } },
              { row: 64, column: 0, input: { type: 'number', value: 202 } },
            ],
          },
        });
      } catch (error) {
        code = (error as { code?: string }).code ?? '';
      }
      return { reached, code, revision: book.revision };
    }, failure);
    expect(result).toEqual({ reached: true, code: 'WORKER_FAILED', revision: initial.revision });
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
      revision: initial.revision,
      values: [10, 20],
      undo: true,
      undone: initial.original,
      redo: true,
      redone: [10, 20],
    });
    await reader.close();
  });

  test(`Worker ${failure} before acknowledgement recovers the durable commit`, async ({
    page,
    context,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const initial = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      const cells = [
        { row: 0, column: 0 },
        { row: 64, column: 0 },
      ];
      const set = (a: number, b: number) =>
        book.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: cells.map((cell, i) => ({
              ...cell,
              input: { type: 'number', value: i ? b : a },
            })),
          },
        });
      await set(10, 20);
      return { id: book.id, revision: book.revision };
    });
    const reader = await context.newPage();
    await reader.goto('/#budget');
    await expect(reader.locator('#take-write-lock')).toBeVisible();
    const result = await page.evaluate(async (kind) => {
      const book = window.opensheet.getWorkbook();
      const worker = (book as any).worker as Worker;
      const oldHandler = worker.onmessage!;
      const oldPost = worker.postMessage.bind(worker);
      let requestId = -1;
      let committedResponse = false;
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'execute') requestId = message.id;
        return Reflect.apply(oldPost, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id !== requestId) return oldHandler.call(worker, event);
        committedResponse = !event.data.error && !!event.data.result?.commit;
        if (kind === 'error')
          worker.onerror?.call(
            worker,
            new ErrorEvent('error', { message: 'Injected after commit' }),
          );
        else worker.onmessageerror?.call(worker, new MessageEvent('messageerror'));
      };
      let code = '';
      try {
        await book.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: book.getSheets()[0].id,
            cells: [
              { row: 0, column: 0, input: { type: 'number', value: 101 } },
              { row: 64, column: 0, input: { type: 'number', value: 202 } },
            ],
          },
        });
      } catch (error) {
        code = (error as { code?: string }).code ?? '';
      }
      return { code, committedResponse, revision: book.revision };
    }, failure);
    expect(result).toEqual({
      code: 'WORKER_FAILED',
      committedResponse: true,
      revision: initial.revision,
    });
    await expect(page.locator('#save-state')).toContainText('Reopen workbook');
    await expect
      .poll(() => reader.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()), {
        timeout: 10000,
      })
      .toBe(true);
    const restored = await reader.evaluate(async () => {
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
    expect(restored).toEqual({
      revision: initial.revision + 1,
      values: [101, 202],
      undo: true,
      undone: [10, 20],
      redo: true,
      redone: [101, 202],
    });
    await reader.close();
  });
}
