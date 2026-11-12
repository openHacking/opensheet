import { expect, test } from '@playwright/test';

test('style toggles read an unloaded cell and preserve other properties through history', async ({
  page,
}) => {
  await page.setViewportSize({ width: 1600, height: 900 });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1').setValues([[42]]);
    await sheet
      .range('A1')
      .setStyle({ bold: true, italic: true, underline: true, wrap: true, color: '#123456' });
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    await app.getGrid()!.ready();
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  for (const [title, property] of [
    ['Bold', 'bold'],
    ['Italic', 'italic'],
    ['Underline', 'underline'],
    ['Wrap text', 'wrap'],
  ] as const) {
    const before = await page.evaluate((title) => {
      const book = window.opensheet.getWorkbook();
      (book as any).cache.clear();
      const button = document.querySelector<HTMLButtonElement>(`button[title="${title}"]`)!;
      button.click();
      return book.revision;
    }, title);
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
      .toBe(before + 1);
    const style = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      return book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
    });
    expect(style[property]).toBe(false);
    expect(style.color).toBe('#123456');
  }
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    await book.undo();
    const sheet = book.getSheets()[0];
    const undone = book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId);
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    const reopened = book.getStyle((await book.getCell(book.getSheets()[0].id, 0, 0))?.styleId);
    return { undone, reopened };
  });
  expect(result.undone.wrap).toBe(true);
  expect(result.reopened).toMatchObject({
    bold: false,
    italic: false,
    underline: false,
    wrap: false,
    color: '#123456',
  });
});

for (const change of ['selection', 'workbook', 'grid'] as const) {
  test(`a delayed style read cannot apply after ${change} changes`, async ({ page }) => {
    await page.setViewportSize({ width: 1600, height: 900 });
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const before = await page.evaluate(async (url) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      await sheet.range('A1').setStyle({ bold: true });
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      await app.getGrid()!.ready();
      (book as any).cache.clear();
      const worker = (book as any).worker as Worker;
      const post = worker.postMessage.bind(worker);
      const handler = worker.onmessage!;
      let readId = -1;
      (window as any).styleCommands = [];
      worker.postMessage = ((message: any, ...rest: any[]) => {
        if (message.method === 'read' && message.args[2] === false && readId === -1)
          readId = message.id;
        if (message.method === 'execute')
          (window as any).styleCommands.push(message.args[0][0].type);
        return Reflect.apply(post, worker, [message, ...rest]);
      }) as typeof worker.postMessage;
      worker.onmessage = (event) => {
        if (event.data.id === readId) (window as any).styleReadHeld = event;
        else handler.call(worker, event);
      };
      (window as any).releaseStyleRead = () => {
        handler.call(worker, (window as any).styleReadHeld);
        worker.postMessage = post;
        worker.onmessage = handler;
      };
      document.querySelector<HTMLButtonElement>('button[title="Bold"]')!.click();
      return book.revision;
    }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
    await page.waitForFunction(() => !!(window as any).styleReadHeld);
    await page.evaluate(async (change) => {
      const app = window.opensheet;
      const sheet = app.getWorkbook().getSheets()[0];
      if (change === 'selection') {
        app.selection.set({
          sheetId: sheet.id,
          startRow: 1,
          endRow: 2,
          startColumn: 0,
          endColumn: 1,
        });
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 1,
        });
      } else if (change === 'workbook') await sheet.range('B1').setValues([[99]]);
      else await app.attachWorkbook(app.getWorkbook());
      (window as any).releaseStyleRead();
    }, change);
    await expect(page.locator('.opensheet')).toHaveAttribute('aria-busy', 'false');
    const result = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      return {
        revision: book.revision,
        commands: (window as any).styleCommands,
        bold: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId).bold,
      };
    });
    expect(result).toEqual({
      revision: before + (change === 'workbook' ? 1 : 0),
      commands: change === 'workbook' ? ['core.cells.set'] : [],
      bold: true,
    });
  });
}
