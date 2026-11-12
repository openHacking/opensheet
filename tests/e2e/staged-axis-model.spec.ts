import { expect, test } from '@playwright/test';

test('a seeded mixed-axis transaction agrees with an independent identity model', async ({
  page,
}) => {
  test.setTimeout(60000);
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 80, columns: 40 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const model = {
      rows: Array.from({ length: 80 }, (_, i) => i),
      columns: Array.from({ length: 40 }, (_, i) => i),
      nextRow: 80,
      nextColumn: 40,
    };
    let seed = 73013;
    const random = () => {
      seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
      return seed >>> 8;
    };
    const coverage = { rowInsert: 0, rowDelete: 0, columnInsert: 0, columnDelete: 0 };
    await book.transaction({}, async (tx) => {
      const staged = tx.getSheets()[0];
      for (let step = 0; step < 100; step++) {
        const row = !!(random() & 1),
          list = row ? model.rows : model.columns,
          insert = list.length < 3 || !!(random() & 1);
        const count = 1 + (random() % 2),
          at = insert ? random() % (list.length + 1) : random() % (list.length - count + 1);
        if (insert) {
          coverage[row ? 'rowInsert' : 'columnInsert']++;
          const next = row ? model.nextRow : model.nextColumn;
          if (row) {
            await staged.insertRows(at, count);
            model.nextRow += count;
          } else {
            await staged.insertColumns(at, count);
            model.nextColumn += count;
          }
          list.splice(at, 0, ...Array.from({ length: count }, (_, i) => next + i));
        } else {
          coverage[row ? 'rowDelete' : 'columnDelete']++;
          if (row) await staged.deleteRows(at, count);
          else await staged.deleteColumns(at, count);
          list.splice(at, count);
        }
        if (staged.rowCount !== model.rows.length || staged.columnCount !== model.columns.length)
          throw new Error(`Wrong staged dimensions at ${step}`);
        if (book.getSheets()[0].rowCount !== 80 || book.getSheets()[0].columnCount !== 40)
          throw new Error('Parent changed during staging');
      }
      await staged
        .range({
          startRow: model.rows.length - 1,
          endRow: model.rows.length,
          startColumn: model.columns.length - 1,
          endColumn: model.columns.length,
        })
        .setValues([[42]]);
    });
    const read = async () => ({
      rows: book.sheetData(sheet.id).rowOrder,
      columns: book.sheetData(sheet.id).columnOrder,
      nextRow: book.sheetData(sheet.id).nextRowId,
      nextColumn: book.sheetData(sheet.id).nextColumnId,
      value: await sheet
        .range({
          startRow: model.rows.length - 1,
          endRow: model.rows.length,
          startColumn: model.columns.length - 1,
          endColumn: model.columns.length,
        })
        .getValues(),
    });
    const committed = await read();
    await book.undo();
    const undo = { rows: sheet.rowCount, columns: sheet.columnCount };
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { model, coverage, committed, undo, reopened: await read() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.committed).toEqual({
    rows: result.model.rows,
    columns: result.model.columns,
    nextRow: result.model.nextRow,
    nextColumn: result.model.nextColumn,
    value: [[42]],
  });
  expect(result.undo).toEqual({ rows: 80, columns: 40 });
  for (const count of Object.values(result.coverage)) expect(count).toBeGreaterThan(0);
  expect(result.reopened).toEqual(result.committed);
});

test('invalid staged axis commands poison and roll back the transaction', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] }));
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[7]]);
    const revision = book.revision,
      bytes = (await book.storageStats()).bytes;
    let caught = '';
    const outcome = await book
      .transaction({}, async (tx) => {
        const staged = tx.getSheets()[0];
        await staged.insertRows(0);
        try {
          await tx.execute({
            type: 'core.axis.insert',
            payload: { sheetId: staged.id, axis: 'row', index: 0, ids: [2] },
          });
        } catch (error) {
          caught = (error as any).code;
        }
      })
      .then(
        () => 'ok',
        (error) => error.code,
      );
    const after = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      rows: sheet.rowCount,
      value: await sheet.range('A1').getValues(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      revision,
      bytes,
      caught,
      outcome,
      after,
      reopened: { rows: sheet.rowCount, value: await sheet.range('A1').getValues() },
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.caught).toBe('INVALID_ARGUMENT');
  expect(result.outcome).toBe('INVALID_ARGUMENT');
  expect(result.after).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    rows: 2,
    value: [[7]],
  });
  expect(result.reopened).toEqual({ rows: 2, value: [[7]] });
});
