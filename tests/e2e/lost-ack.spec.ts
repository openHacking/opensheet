import { expect, test } from '@playwright/test';

test('lost acknowledgement reopens the committed revision and rejects a stale retry', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const values = await sheet.range('A1:B1').getValues();
    const command = {
      protocolVersion: 1 as const,
      workbookId: book.id,
      commandId: `lost-ack-${crypto.randomUUID()}`,
      baseRevision: book.revision,
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [
          { row: 0, column: 0, input: { type: 'number', value: 111 } },
          { row: 0, column: 1, input: { type: 'number', value: 222 } },
        ],
      },
    };
    const worker = (book as any).worker as Worker;
    const handler = worker.onmessage!;
    const post = worker.postMessage.bind(worker);
    let executeId = -1;
    (book as any).options.operationTimeoutMs = 5000;
    (window as any).lostAck = false;
    (window as any).lostCommand = command;
    (window as any).commitEvents = 0;
    book.onCommit(() => (window as any).commitEvents++);
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') executeId = message.id;
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (event.data.id === executeId) {
        (window as any).lostAck = !event.data.error && !!event.data.result?.commit;
      } else handler.call(worker, event);
    };
    (window as any).lostResult = book.execute(command).then(
      () => 'unexpected success',
      (error: { code?: string }) => error.code,
    );
    return { id: book.id, revision: book.revision, values };
  });
  await page.waitForFunction(() => (window as any).lostAck, null, { timeout: 4000 });
  await expect(page.locator('#save-state')).toHaveText('Saving…');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before.revision);
  expect(await page.evaluate(() => (window as any).commitEvents)).toBe(0);
  expect(await page.evaluate(() => (window as any).lostResult)).toBe('WORKER_TIMEOUT');
  await expect(page.locator('#save-state')).toHaveText('Save unconfirmed · Reopen workbook');
  const restored = await page.evaluate(async (id) => {
    await window.opensheet.open(id);
    const book = window.opensheet.getWorkbook();
    let retry = '';
    try {
      await book.execute((window as any).lostCommand);
    } catch (error) {
      retry = (error as { code?: string }).code ?? '';
    }
    const revisionAfterRetry = book.revision;
    const values = await book.getSheets()[0].range('A1:B1').getValues();
    const undo = await book.undo();
    const afterUndo = await book.getSheets()[0].range('A1:B1').getValues();
    const redo = await book.redo();
    const afterRedo = await book.getSheets()[0].range('A1:B1').getValues();
    return { retry, revisionAfterRetry, values, undo, afterUndo, redo, afterRedo };
  }, before.id);
  expect(restored).toEqual({
    retry: 'REVISION_CONFLICT',
    revisionAfterRetry: before.revision + 1,
    values: [[111, 222]],
    undo: true,
    afterUndo: before.values,
    redo: true,
    afterRedo: [[111, 222]],
  });
  await page.evaluate(async (id) => window.opensheet.open(id), before.id);
  expect(
    await page.evaluate(async () =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
    ),
  ).toEqual([[111, 222]]);
});
