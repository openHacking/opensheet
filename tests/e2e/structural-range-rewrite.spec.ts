import { expect, test } from '@playwright/test';

test('partial range deletion keeps surviving cells through history and reopen', async ({
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
          { name: 'Source', rows: 8, columns: 3 },
          { name: 'Other', rows: 3, columns: 5 },
        ],
      }),
    );
    let book = app.getWorkbook();
    let [source, other] = book.getSheets();
    await source.range('A1:A5').setValues([[1], [2], [3], [4], [5]]);
    await other
      .range('A1:D1')
      .setFormulas([
        ['SUM(Source!A1:A5)', 'SUM(Source!A5:A2)', 'SUM(Source!A1:A2)', 'SUM(Source!A2:A4)'],
      ]);
    const read = async () => ({
      formulas: await other.range('A1:D1').getFormulas(),
      values: await other.range('A1:D1').getValues(),
    });
    const before = await read();
    await source.deleteRows(0, 2);
    const after = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    const reopened = await read();
    await source.range('B1:C1').setValues([[7, 8]]);
    await source.range('C2').setFormulas([['SUM(A1:C1)']]);
    await other
      .range('A2:C2')
      .setFormulas([['SUM(Source!A1:C1)', 'SUM(Source!C1:A1)', 'Source!A1']]);
    const readColumns = async (local: string) => ({
      other: await other.range('A2:C2').getValues(),
      local: await source.range(local).getValues(),
      formula: await source.range(local).getFormulas(),
    });
    const columnsBefore = await readColumns('C2');
    await source.deleteColumns(0);
    const columnsAfter = await readColumns('B2');
    await book.undo();
    const columnsUndo = await readColumns('C2');
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    [source, other] = book.getSheets();
    return {
      before,
      after,
      undone,
      reopened,
      columnsBefore,
      columnsAfter,
      columnsUndo,
      columnsReopened: await readColumns('B2'),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.before.values).toEqual([[15, 14, 3, 9]]);
  expect(result.after.values).toEqual([[12, 12, { error: '#REF!' }, 7]]);
  expect(result.undone).toEqual(result.before);
  expect(result.reopened).toEqual(result.after);
  expect(result.columnsBefore.other).toEqual([[18, 18, 3]]);
  expect(result.columnsBefore.local).toEqual([[18]]);
  expect(result.columnsAfter.other).toEqual([[15, 15, { error: '#REF!' }]]);
  expect(result.columnsAfter.local).toEqual([[15]]);
  expect(result.columnsUndo).toEqual(result.columnsBefore);
  expect(result.columnsReopened).toEqual(result.columnsAfter);
});
