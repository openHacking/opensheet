import { expect, test } from '@playwright/test';

test('deleting a row containing an invalid formula preserves other formulas and history', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 4, columns: 2 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([[1], [2], [3]]);
    await sheet.range('B1:B2').setFormulas([['A1+'], ['A2*2']]);
    const before = { formula: await sheet.range('B1:B2').getFormulas(), revision: book.revision };
    const outcome = await sheet.deleteRows(0).then(
      () => 'ok',
      (error) => error.code ?? String(error),
    );
    if (outcome !== 'ok') return { outcome, before };
    const after = {
      formula: await sheet.range('B1').getFormulas(),
      value: await sheet.range('B1').getValues(),
    };
    await book.undo();
    const undo = await sheet.range('B1:B2').getFormulas();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = {
      formula: await sheet.range('B1').getFormulas(),
      value: await sheet.range('B1').getValues(),
    };
    await sheet.range('B2').setFormulas([['A1+']]);
    const retainedRevision = book.revision;
    const retainedBefore = await sheet.range('B1:B2').getFormulas();
    const retainedOutcome = await sheet.deleteRows(0).then(
      () => 'ok',
      (error) => error.code,
    );
    const retainedAfter = await sheet.range('B1:B2').getFormulas();
    return {
      outcome,
      before,
      after,
      undo,
      reopened,
      retainedRevision,
      retainedAfterRevision: book.revision,
      retainedBefore,
      retainedAfter,
      retainedOutcome,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.outcome).toBe('ok');
  expect(result.after?.value).toEqual([[4]]);
  expect(result.after?.formula[0][0]).toContain('A1');
  expect(result.undo).toEqual(result.before.formula);
  expect(result.reopened).toEqual(result.after);
  expect(result.retainedOutcome).toBe('UNSUPPORTED_FEATURE');
  expect(result.retainedAfter).toEqual(result.retainedBefore);
  expect(result.retainedAfterRevision).toBe(result.retainedRevision);
});

test('deleting a column or sheet can discard malformed formulas without rewriting them', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(
      createWorkbookFile({
        sheets: [
          { name: 'Data', rows: 3, columns: 3 },
          { name: 'Other', rows: 3, columns: 2 },
        ],
      }),
    );
    let book = app.getWorkbook();
    let [data, other] = book.getSheets();
    await data.range('A1:B1').setFormulas([['A1+', 'B2*2']]);
    await data.range('B2').setValues([[3]]);
    await other.range('A1').setFormulas([['Data!B2']]);
    await data.deleteColumns(0);
    const afterColumn = {
      local: await data.range('A1').getValues(),
      external: await other.range('A1').getValues(),
    };
    await book.undo();
    const undo = await data.range('A1:B1').getFormulas();
    await book.redo();
    await other.range('B1').setFormulas([['A1+']]);
    await book.removeSheet(other.id);
    const afterSheet = book.getSheets().map((sheet) => sheet.name);
    await book.undo();
    const restored = await book.getSheets()[1].range('B1').getFormulas();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return {
      afterColumn,
      undo,
      afterSheet,
      restored,
      reopened: book.getSheets().map((sheet) => sheet.name),
      value: await book.getSheets()[0].range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.afterColumn).toEqual({ local: [[6]], external: [[3]] });
  expect(result.undo).toEqual([['A1+', 'B2*2']]);
  expect(result.afterSheet).toEqual(['Data']);
  expect(result.restored).toEqual([['A1+']]);
  expect(result.reopened).toEqual(['Data']);
  expect(result.value).toEqual([[6]]);
});

test('a late surviving malformed formula rolls back staged deletion pages', async ({ page }) => {
  test.setTimeout(60_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
      let armed = false;
      addEventListener('message', ({data}) => { if (data.watchDeletePage) armed = true; });
      const original = IDBObjectStore.prototype.put;
      IDBObjectStore.prototype.put = function(...args) {
        const request = original.apply(this, args);
        if (armed && this.name === 'records') this.transaction.addEventListener('complete', () => {
          armed = false; postMessage({ deletePageStaged: true });
        }, {once:true});
        return request;
      };
    `,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 1 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [
          { row: 0, column: 0, input: { type: 'formula', expression: 'A1+' } },
          { row: 1, column: 0, input: { type: 'formula', expression: 'A130' } },
          { row: 64, column: 0, input: { type: 'formula', expression: 'A1+' } },
          { row: 129, column: 0, input: { type: 'number', value: 17 } },
        ],
      },
    });
    const request = <T>(r: IDBRequest<T>) =>
      new Promise<T>((resolve, reject) => {
        r.onsuccess = () => resolve(r.result);
        r.onerror = () => reject(r.error);
      });
    const raw = async () => {
      const db = await request(indexedDB.open('opensheet'));
      try {
        const tx = db.transaction(['heads', 'records', 'sizes', 'scratch']);
        const keys = IDBKeyRange.bound(`${book.id}/`, `${book.id}/\uffff`);
        return await Promise.all([
          request(tx.objectStore('heads').get(book.id)),
          ...['records', 'sizes', 'scratch'].map((name) =>
            request(tx.objectStore(name).getAllKeys(keys)),
          ),
        ]);
      } finally {
        db.close();
      }
    };
    const baseline = await raw(),
      revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    let staged = false;
    const worker = (book as any).worker as Worker;
    worker.addEventListener('message', ({ data }) => {
      if (data.deletePageStaged) staged = true;
    });
    worker.postMessage({ watchDeletePage: true });
    const outcome = await sheet.deleteRows(0).then(
      () => 'ok',
      (error) => error.code,
    );
    const afterRaw = await raw();
    const after = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      first: await sheet.range('A1:A2').getFormulas(),
      late: await sheet.range('A65').getFormulas(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = {
      first: await sheet.range('A1:A2').getFormulas(),
      late: await sheet.range('A65').getFormulas(),
    };
    await sheet.range('A65').clear();
    await sheet.deleteRows(0);
    const retry = {
      formula: await sheet.range('A1').getFormulas(),
      value: await sheet.range('A1').getValues(),
    };
    await book.undo();
    const undo = await sheet.range('A1:A2').getFormulas();
    await book.redo();
    return {
      staged,
      outcome,
      baseline,
      afterRaw,
      revision,
      bytes,
      after,
      reopened,
      retry,
      undo,
      redo: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.staged).toBe(true);
  expect(result.outcome).toBe('UNSUPPORTED_FEATURE');
  expect(result.afterRaw).toEqual(result.baseline);
  expect(result.after).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    first: [['A1+'], ['A130']],
    late: [['A1+']],
  });
  expect(result.reopened).toEqual({ first: result.after.first, late: result.after.late });
  expect(result.retry.value).toEqual([[17]]);
  expect(result.retry.formula[0][0]).toContain('A129');
  expect(result.undo).toEqual([['A1+'], ['A130']]);
  expect(result.redo).toEqual([[17]]);
});
