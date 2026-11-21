import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('order workload preserves inputs, calculations and metadata through native exports', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      let book = app.getWorkbook();
      let [orders, summary] = book.getSheets();
      const ids = [orders.id, summary.id];
      await orders.range('A1:D4').setValues([
        ['ID', 'Quantity', 'Price', 'Active'],
        ['0001', 3, 12.5, true],
        ['0002', 2, 7.25, false],
        ['0003', 1, 100, true],
      ]);
      await orders.range('E1').setValues([['Total']]);
      await orders.range('E2:E4').setFormulas([['B2*C2'], ['B3*C3'], ['B4*C4']]);
      await summary.range('A1').setFormulas([['SUM(Orders!E2:E4)']]);
      const before = await summary.range('A1').getValues();
      await orders.range('B2').setValues([[4]]);
      const edited = await summary.range('A1').getValues();
      await book.undo();
      const undone = await summary.range('A1').getValues();
      await book.redo();
      const redone = await summary.range('A1').getValues();
      await orders.setFilter({
        mode: 'all',
        headerRows: 1,
        criteria: [{ column: 3, operator: 'eq', value: true }],
      });
      const hidden = book.filteredRows(orders.id);
      await orders.setFilter(null);
      await orders.sort('A2:E4', 0, 'desc');
      await orders.range('C2:C4').setStyle({ numberFormat: '0.00' });
      await orders.range('G1').setValues([['Orders']]);
      await orders.range('G1:H1').merge();
      const inspect = async () => ({
        ids: book.getSheets().map((sheet) => sheet.id),
        values: await orders.range('A2:E4').getValues(),
        formulas: await orders.range('E2:E4').getFormulas(),
        total: await summary.range('A1').getValues(),
        summaryFormula: await summary.range('A1').getFormulas(),
        styles: await Promise.all(
          [1, 2, 3].map(async (row) =>
            book.getStyle((await book.getCell(orders.id, row, 2))?.styleId),
          ),
        ),
        merges: book.sheetData(orders.id).merges,
        title: await orders.range('G1').getValues(),
      });
      const original = await inspect();
      const json = await book.toJSON();
    const chunks: BlobPart[] = [];
    for await (const chunk of book.streamBinary()) chunks.push(Uint8Array.from(chunk));
      const binary = new Blob(chunks);
      await book.importJSON(json);
      [orders, summary] = book.getSheets();
      const afterJSON = await inspect();
      await book.importBinary(binary);
      [orders, summary] = book.getSheets();
      await app.open(book.id);
      book = app.getWorkbook();
      [orders, summary] = book.getSheets();
      return {
        ids,
        before,
        edited,
        undone,
        redone,
        hidden,
        original,
        afterJSON,
        afterBinary: await inspect(),
      };
    },
    JSON.stringify(
      createWorkbookFile({
        sheets: [
          { name: 'Orders', rows: 10, columns: 8 },
          { name: 'Summary', rows: 10, columns: 8 },
        ],
      }),
    ),
  );

  expect(result.before).toEqual([[152]]);
  expect(result.edited).toEqual([[164.5]]);
  expect(result.undone).toEqual([[152]]);
  expect(result.redone).toEqual([[164.5]]);
  expect(result.hidden).toEqual([2, 4, 5, 6, 7, 8, 9]);
  expect(result.original.ids).toEqual(result.ids);
  expect(result.original.values).toEqual([
    ['0003', 1, 100, true, 100],
    ['0002', 2, 7.25, false, 14.5],
    ['0001', 4, 12.5, true, 50],
  ]);
  expect(result.original.total).toEqual([[164.5]]);
  expect(result.original.formulas.flat().every((formula) => typeof formula === 'string')).toBe(
    true,
  );
  expect(result.original.summaryFormula).toEqual([['SUM(Orders!E2:E4)']]);
  expect(result.original.title).toEqual([['Orders']]);
  expect(result.original.merges).toEqual([
    { startRow: 0, startColumn: 6, endRow: 1, endColumn: 8 },
  ]);
  expect(result.original.styles).toEqual(
    Array.from({ length: 3 }, () => ({ numberFormat: '0.00' })),
  );
  expect(result.afterJSON).toEqual(result.original);
  expect(result.afterBinary).toEqual(result.original);
});
