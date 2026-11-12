import { expect, test } from '@playwright/test';

test('column deletion retains reordered IDs, partial pages and cross-sheet references', async ({
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
          { name: 'Source', rows: 2, columns: 70 },
          { name: 'Other', rows: 2, columns: 2 },
        ],
      }),
    );
    let book = app.getWorkbook();
    let [source, other] = book.getSheets();
    await source
      .range({ startRow: 0, endRow: 1, startColumn: 0, endColumn: 70 })
      .fillSeries({ start: 1, step: 1, axis: 'column' });
    await source.insertColumns(0);
    await source.range('A1').setValues([[999]]);
    await source.setColumnWidth(41, 125);
    await other.range('A1:B1').setFormulas([['Source!B1', 'Source!AP1']]);
    const read = async () => {
      const data = book.sheetData(source.id);
      return {
        ids: [...data.columnOrder],
        width: data.columns[40]?.size,
        values: await source
          .range({ startRow: 0, endRow: 1, startColumn: 0, endColumn: data.columnOrder.length })
          .getValues(),
        formulas: await other.range('A1:B1').getFormulas(),
        results: await other.range('A1:B1').getValues(),
      };
    };
    const before = await read();
    await source.deleteColumns(1, 40);
    const deleted = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    const reopened = await read();
    await source.deleteColumns(25, 6);
    const sharedBucket = await read();
    await book.undo();
    const sharedUndo = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    return {
      before,
      deleted,
      undone,
      reopened,
      sharedBucket,
      sharedUndo,
      sharedReopened: await read(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.before.ids).toEqual([70, ...Array.from({ length: 70 }, (_, i) => i)]);
  expect(result.before.values).toEqual([[999, ...Array.from({ length: 70 }, (_, i) => i + 1)]]);
  expect(result.before.results).toEqual([[1, 41]]);
  expect(result.deleted.ids).toEqual([70, ...Array.from({ length: 30 }, (_, i) => i + 40)]);
  expect(result.deleted.values).toEqual([[999, ...Array.from({ length: 30 }, (_, i) => i + 41)]]);
  expect(result.deleted.width).toBe(125);
  expect(result.deleted.results).toEqual([[{ error: '#REF!' }, 41]]);
  expect(result.undone).toEqual(result.before);
  expect(result.reopened).toEqual(result.deleted);
  expect(result.sharedBucket.ids).toEqual([70, ...Array.from({ length: 24 }, (_, i) => i + 40)]);
  expect(result.sharedBucket.values).toEqual([
    [999, ...Array.from({ length: 24 }, (_, i) => i + 41)],
  ]);
  expect(result.sharedBucket.results).toEqual([[{ error: '#REF!' }, 41]]);
  expect(result.sharedUndo).toEqual(result.deleted);
  expect(result.sharedReopened).toEqual(result.sharedBucket);
});
