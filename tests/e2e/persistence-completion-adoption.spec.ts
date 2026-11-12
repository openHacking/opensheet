import { expect, test } from '@playwright/test';

for (const operation of ['undo', 'redo', 'generate', 'json', 'binary'] as const) {
  test(`${operation} announces saved only after adopting its result`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(
      async ({ url, operation }) => {
        const { createWorkbookFile } = await import(/* @vite-ignore */ url);
        const app = window.opensheet;
        await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
        const book = app.getWorkbook();
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([[7]]);
        const json = await book.toJSON();
        const chunks: BlobPart[] = [];
        for await (const chunk of book.streamBinary()) chunks.push(new Uint8Array(chunk).buffer);
        const binary = new Blob(chunks);
        await book
          .getSheets()[0]
          .range('A1')
          .setValues([[9]]);
        if (operation === 'redo') await book.undo();
        const previous = book.revision;
        const view = () => ({
          revision: book.revision,
          sheets: book
            .getSheets()
            .map((sheet) => ({ id: sheet.id, rows: sheet.rowCount, columns: sheet.columnCount })),
        });
        let saving = false;
        let observed: ReturnType<typeof view> | undefined;
        const stop = book.onPersistence((state) => {
          if (state.state === 'saving') saving = true;
          if (saving && state.state === 'saved' && !observed) observed = view();
        });
        if (operation === 'undo') await book.undo();
        else if (operation === 'redo') await book.redo();
        else if (operation === 'generate') await book.generate(100);
        else if (operation === 'json') await book.importJSON(json);
        else await book.importBinary(binary);
        stop();
        return { previous, observed, current: view() };
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, operation },
    );
    expect(result.current.revision).toBeGreaterThan(result.previous);
    expect(result.observed).toEqual(result.current);
  });
}

test('commit observers see the adopted revision while acknowledgement is still pending', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    const before = book.revision;
    let observed: { revision: number; pending: number; rows: number } | undefined;
    let copy: Promise<string> | undefined;
    const stop = book.onCommit(() => {
      if (observed) return;
      observed = {
        revision: book.revision,
        pending: book.persistenceState.pending,
        rows: book.sheetData(sheet.id).rowOrder.length,
      };
      copy = sheet
        .range('A2')
        .copyTo('B2')
        .then(
          () => 'success',
          (error) => error.code,
        );
    });
    await sheet.insertRows(0);
    const code = await copy;
    stop();
    const saved = book.persistenceState;
    await sheet.range('A2').copyTo('B2');
    return { before, observed, code, saved: { state: saved.state, pending: saved.pending } };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.observed).toEqual({ revision: result.before + 1, pending: 1, rows: 4 });
  expect(result.code).toBe('BUSY');
  expect(result.saved).toEqual({ state: 'saved', pending: 0 });
});
