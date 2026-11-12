import { expect, test } from '@playwright/test';

test('persistence observer writes are queued after the operation that notified them', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[5, 9]]);
    let triggered = false;
    let insertion: Promise<unknown> | undefined;
    const stop = book.onPersistence((state) => {
      if (!triggered && state.state === 'saving') {
        triggered = true;
        insertion = sheet.insertRows(0);
      }
    });
    const copy = sheet.range('A1').copyTo('B1');
    await copy;
    await insertion;
    stop();
    const current = await sheet.range('A1:B2').getValues();
    await book.undo();
    const undoInsert = await sheet.range('A1:B1').getValues();
    await book.undo();
    const undoCopy = await sheet.range('A1:B1').getValues();
    await book.redo();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      triggered,
      current,
      undoInsert,
      undoCopy,
      reopened: await sheet.range('A1:B2').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    triggered: true,
    current: [
      [null, null],
      [5, 5],
    ],
    undoInsert: [[5, 5]],
    undoCopy: [[5, 9]],
    reopened: [
      [null, null],
      [5, 5],
    ],
  });
});

test('persistence observers cannot change a command already being submitted', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 2, columns: 2 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    const payload = {
      sheetId: sheet.id,
      cells: [{ row: 0, column: 0, input: { type: 'number', value: 7 } }],
    };
    let changed = false;
    const stop = book.onPersistence((state) => {
      if (state.state === 'saving') {
        changed = true;
        payload.cells[0].input.value = 99;
        payload.cells[0].column = 1;
      }
    });
    await book.execute({ type: 'core.cells.set', payload });
    stop();
    return { changed, values: await sheet.range('A1:B1').getValues() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ changed: true, values: [[7, null]] });
});

test('observer mutation cannot change the remembered idempotency key', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 2, columns: 2 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    const command = {
      protocolVersion: 1 as const,
      commandId: 'original-id',
      workbookId: book.id,
      baseRevision: book.revision,
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 7 } }],
      },
    };
    const original = structuredClone(command);
    const stop = book.onPersistence((state) => {
      if (state.state === 'saving') command.commandId = 'modified-id';
    });
    await book.execute(command);
    stop();
    const revision = book.revision;
    const duplicate = await book.execute(original);
    return {
      revision,
      after: book.revision,
      duplicateRevision: duplicate?.revision,
      values: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.after).toBe(result.revision);
  expect(result.duplicateRevision).toBe(result.revision);
  expect(result.values).toEqual([[7]]);
});

test('saved notification exposes the newly adopted structural revision', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[1, 9]]);
    const priorRevision = book.revision;
    let sawSaving = false;
    let observed: { revision: number; rows: number } | undefined;
    let followup: Promise<unknown> | undefined;
    const stop = book.onPersistence((state) => {
      if (state.state === 'saving') sawSaving = true;
      if (sawSaving && state.state === 'saved' && !observed) {
        observed = { revision: book.revision, rows: book.sheetData(sheet.id).rowOrder.length };
        followup = sheet.range('A2').copyTo('B2');
      }
    });
    await sheet.insertRows(0);
    await followup;
    stop();
    return { priorRevision, observed, values: await sheet.range('A1:B2').getValues() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.observed).toEqual({ revision: result.priorRevision + 1, rows: 4 });
  expect(result.values).toEqual([
    [null, null],
    [1, 1],
  ]);
});
