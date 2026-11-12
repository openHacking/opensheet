import { expect, test } from '@playwright/test';

test('copy and fill reject while a structural write is awaiting acknowledgement', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 4, columns: 2 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:B2').setValues([
      [1, 9],
      [2, 8],
    ]);
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    let firstId = -1;
    let held!: MessageEvent;
    let reached!: () => void;
    const waiting = new Promise<void>((resolve) => {
      reached = resolve;
    });
    const submitted: string[] = [];
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') {
        submitted.push(message.args[0][0].type);
        if (firstId === -1) firstId = message.id;
      }
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (event.data.id === firstId) {
        held = event;
        reached();
      } else handler.call(worker, event);
    };
    const insertion = sheet.insertRows(0);
    await waiting;
    const copy = await sheet
      .range('A1')
      .copyTo('B1')
      .then(
        () => 'success',
        (error) => error.code,
      );
    const fill = await sheet
      .range('A1:A2')
      .fillDown()
      .then(
        () => 'success',
        (error) => error.code,
      );
    handler.call(worker, held);
    worker.postMessage = post;
    worker.onmessage = handler;
    await insertion;
    const after = await sheet.range('A1:B3').getValues();
    await sheet.range('A2').copyTo('B2');
    await sheet.range('A2:A3').fillDown();
    const retried = await sheet.range('A1:B3').getValues();
    await book.undo();
    await book.undo();
    const undone = await sheet.range('A1:B3').getValues();
    await book.redo();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      copy,
      fill,
      submitted,
      after,
      retried,
      undone,
      reopened: await sheet.range('A1:B3').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    copy: 'BUSY',
    fill: 'BUSY',
    submitted: ['core.axis.insert'],
    after: [
      [null, null],
      [1, 9],
      [2, 8],
    ],
    retried: [
      [null, null],
      [1, 1],
      [1, 8],
    ],
    undone: [
      [null, null],
      [1, 9],
      [2, 8],
    ],
    reopened: [
      [null, null],
      [1, 1],
      [1, 8],
    ],
  });
});
