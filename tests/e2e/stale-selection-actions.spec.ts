import { expect, test } from '@playwright/test';

test('keyboard, toolbar and paste reject stale selections after a pending structural edit', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1600, height: 900 });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A2').setValues([[1], [2]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
    const grid = app.getGrid() as any;
    (window as any).selectionErrors = [];
    grid.options.onError = (error: { code?: string }) =>
      (window as any).selectionErrors.push(error.code);
    const worker = (book as any).worker as Worker;
    const post = worker.postMessage.bind(worker);
    const handler = worker.onmessage!;
    let firstId = -1;
    (window as any).selectionCommands = [];
    worker.postMessage = ((message: any, ...rest: any[]) => {
      if (message.method === 'execute') {
        (window as any).selectionCommands.push(message.args[0][0].type);
        if (firstId === -1) firstId = message.id;
      }
      return Reflect.apply(post, worker, [message, ...rest]);
    }) as typeof worker.postMessage;
    worker.onmessage = (event) => {
      if (event.data.id === firstId) (window as any).selectionHeld = event;
      else handler.call(worker, event);
    };
    (window as any).selectionRelease = () => {
      handler.call(worker, (window as any).selectionHeld);
      worker.postMessage = post;
      worker.onmessage = handler;
    };
    (window as any).selectionInsert = sheet.insertRows(0);
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  await page.waitForFunction(() => !!(window as any).selectionHeld);
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await page.keyboard.press('Delete');
  await page.keyboard.press('Control+d');
  await expect
    .poll(() => page.evaluate(() => (window as any).selectionErrors))
    .toEqual(['BUSY', 'BUSY']);
  await page.getByRole('button', { name: 'Bold', exact: true }).click();
  await page.evaluate(() => window.opensheet.getGrid()!.pasteText('9'));
  await expect
    .poll(() => page.evaluate(() => (window as any).selectionErrors))
    .toEqual(['BUSY', 'BUSY', 'REVISION_CONFLICT']);
  expect(await page.evaluate(() => (window as any).selectionCommands)).toEqual([
    'core.axis.insert',
    'core.cells.pasteTile',
  ]);
  await page.evaluate(async () => {
    (window as any).selectionRelease();
    await (window as any).selectionInsert;
  });
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    return {
      values: await sheet.range('A1:A3').getValues(),
      style: book.getStyle((await book.getCell(sheet.id, 1, 0))?.styleId),
    };
  });
  expect(result.values).toEqual([[null], [1], [2]]);
  expect(result.style.bold).not.toBe(true);
});

test('a toolbar action cannot move to a different selection before it runs', async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 900 });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    const revision = book.revision;
    const execute = book.execute.bind(book);
    let writes = 0;
    (book as any).execute = (...args: unknown[]) => {
      writes++;
      return (execute as any)(...args);
    };
    const button = document.querySelector<HTMLButtonElement>('button[title="Bold"]')!;
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    button.click();
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 });
    await Promise.resolve();
    return { revision, after: book.revision, pending: book.persistenceState.pending, writes };
  });
  expect(result.after).toBe(result.revision);
  expect(result.pending).toBe(0);
  expect(result.writes).toBe(0);
});
