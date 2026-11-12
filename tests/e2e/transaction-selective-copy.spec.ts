import { expect, test } from '@playwright/test';

test('selective copy reads staged cells and preserves overlap, history and reopen', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 4, columns: 5 }] }));
    let book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:C1').setValues([[1, 2, 3]]);
    await book.transaction({}, async (tx) => {
      const staged = tx.getSheets()[0];
      await staged.range('A1').setValues([[9]]);
      await staged.range('B1').setFormulas([['=A1*2']]);
      await staged.range('A1:B1').copyTo('B1:C1', { mode: 'values' });
      await staged.range('A2').setFormulas([['=$A1+B1']]);
      await staged.range('A2').copyTo('B2', { mode: 'formulas' });
      await staged.range('A3').setStyle({ bold: true });
      await staged.range('A3').copyTo('B3', { mode: 'formats' });
    });
    const read = async () => ({
      values: await book.getSheets()[0].range('A1:C1').getValues(),
      formula: await book.getSheets()[0].range('B2').getFormulas(),
      style: (
        await book.readCells(
          book.getSheets()[0].id,
          { startRow: 2, endRow: 3, startColumn: 1, endColumn: 2 },
          { fields: ['style'] },
        )
      ).cells[0].style,
    });
    const committed = await read();
    await book.undo();
    const undo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.committed.values).toEqual([[9, 9, 18]]);
  expect(result.committed.formula).toEqual([['($A1+C1)']]);
  expect(result.committed.style).toMatchObject({ bold: true });
  expect(result.undo.values).toEqual([[1, 2, 3]]);
  expect(result.undo.formula).toEqual([[null]]);
  expect(result.undo.style).toEqual({});
  expect(result.reopened).toEqual(result.committed);
});

test('selective copy supports new sheets and rolls back on merged targets', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 4, columns: 4 }] }));
    let book = app.getWorkbook();
    const revision = book.revision;
    const failure = await book
      .transaction({}, async (tx) => {
        const sheet = tx.getSheets()[0];
        await sheet.range('A1').setValues([[42]]);
        await sheet.range('C1:D1').merge();
        await sheet.range('A1').copyTo('C1', { mode: 'values' });
      })
      .then(
        () => 'ok',
        (error) => error.code,
      );
    const afterFailure = {
      revision: book.revision,
      value: await book.getSheets()[0].range('A1').getValues(),
      merges: book.sheetData(book.getSheets()[0].id).merges.length,
    };
    await book.transaction({}, async (tx) => {
      const sheet = await tx.addSheet('New', { rows: 1, columns: 1 });
      await sheet.insertColumns(1);
      await sheet.range('A1').setValues([[17]]);
      await sheet.range('A1').copyTo('B1', { mode: 'values' });
    });
    await app.open(book.id);
    book = app.getWorkbook();
    return {
      revision,
      failure,
      afterFailure,
      value: await book.getSheetByName('New')!.range('A1:B1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.failure).toBe('INVALID_RANGE');
  expect(result.afterFailure).toEqual({ revision: result.revision, value: [[null]], merges: 0 });
  expect(result.value).toEqual([[17, 17]]);
});

test('staged values copy preserves scalars, blanks, errors and cross-sheet formulas', async ({
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
          { name: 'Data', rows: 2, columns: 6 },
          { name: 'Other', rows: 1, columns: 1 },
        ],
      }),
    );
    const book = app.getWorkbook();
    await book.transaction({}, async (tx) => {
      const sheet = tx.getSheetByName('Data')!;
      await tx
        .getSheetByName('Other')!
        .range('A1')
        .setValues([[7]]);
      await sheet.range('A1:D1').setValues([[0, false, 'text', null]]);
      await tx.execute({
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [{ row: 0, column: 4, input: { type: 'error', code: '#DIV/0!' } }],
        },
      });
      await sheet.range('F1').setFormulas([['=Other!A1*2']]);
      await sheet.range('A1:F1').copyTo('A2:F2', { mode: 'values' });
    });
    const sheet = book.getSheetByName('Data')!;
    return {
      values: await sheet.range('A2:F2').getValues(),
      formulas: await sheet.range('A2:F2').getFormulas(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.values).toEqual([[0, false, 'text', null, { error: '#DIV/0!' }, 14]]);
  expect(result.formulas).toEqual([[null, null, null, null, null, null]]);
});

test('cancelling a large selective copy rolls back earlier staged writes', async ({ page }) => {
  test.setTimeout(60000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let watchSelective = false, selectiveCount = 0;
addEventListener('message', ({data}) => { if (data.watchSelective) watchSelective = true; });
const selectiveStringify = JSON.stringify;
JSON.stringify = function(value, ...rest) {
  if (watchSelective && value && Number.isInteger(value.row) && Number.isInteger(value.column) && value.input) {
    if (++selectiveCount === 128) { watchSelective = false; postMessage({ selectiveStarted: true }); }
  }
  return selectiveStringify.call(this, value, ...rest);
};`,
    });
  });
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 1000, columns: 200 }] }));
    let book = app.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[1]]);
    const revision = book.revision;
    const bytes = (await book.storageStats()).bytes;
    const worker = (book as any).worker as Worker;
    let reached = false;
    worker.addEventListener('message', ({ data }) => {
      if (data.selectiveStarted) {
        reached = true;
        book.cancel();
      }
    });
    worker.postMessage({ watchSelective: true });
    const code = await book
      .transaction({}, async (tx) => {
        const sheet = tx.getSheets()[0];
        await sheet.range('A1').setValues([[9]]);
        await sheet.range('A1:CV1000').copyTo('CW1:GR1000', { mode: 'values' });
      })
      .then(
        () => 'ok',
        (error) => error.code,
      );
    const after = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      value: await book.getSheets()[0].range('A1').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    const reopened = await book.getSheets()[0].range('A1').getValues();
    await book.transaction({}, async (tx) => {
      const sheet = tx.getSheets()[0];
      await sheet.range('A1').setValues([[9]]);
      await sheet.range('A1:CV1000').copyTo('CW1:GR1000', { mode: 'values' });
    });
    const copied = await book.getSheets()[0].range('CW1').getValues();
    await book.undo();
    const undone = await book.getSheets()[0].range('A1').getValues();
    await book.redo();
    const redone = await book.getSheets()[0].range('CW1').getValues();
    return {
      reached,
      code,
      revision,
      bytes,
      after,
      reopened,
      copied,
      undone,
      redone,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.reached).toBe(true);
  expect(result.code).toBe('ABORTED');
  expect(result.after).toEqual({ revision: result.revision, bytes: result.bytes, value: [[1]] });
  expect(result.reopened).toEqual([[1]]);
  expect(result.copied).toEqual([[9]]);
  expect(result.undone).toEqual([[1]]);
  expect(result.redone).toEqual([[9]]);
});
