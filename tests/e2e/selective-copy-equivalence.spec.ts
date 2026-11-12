import { expect, test } from '@playwright/test';

test('transaction and direct selective copies agree for overlapping mixed cells', async ({
  page,
}) => {
  test.setTimeout(60000);
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    const results = [];
    for (const mode of ['values', 'formulas', 'formats'] as const) {
      await app.load(
        createWorkbookFile({
          sheets: [
            { name: 'Direct', rows: 5, columns: 4 },
            { name: 'Staged', rows: 5, columns: 4 },
          ],
        }),
      );
      const book = app.getWorkbook();
      for (const sheet of book.getSheets()) {
        await sheet.range('A1:C3').setValues([
          [1, 'text', null],
          [false, 3, 4],
          [7, 8, 9],
        ]);
        await sheet.range('B1').setFormulas([['=$A1+C2']]);
        await sheet.range('A1').setStyle({ bold: true, numberFormat: '0.00' });
        await sheet.range('B2').setStyle({ italic: true });
        await sheet.range('B2').setNote('keep note');
        await sheet.range('C3').setLink({ target: 'https://example.com', tooltip: 'link' });
      }
      const [direct, staged] = book.getSheets();
      await direct.range('A1:C3').copyTo('B2:D4', { mode });
      await book.transaction({}, async (tx) => {
        await tx.getSheetById(staged.id)!.range('A1:C3').copyTo('B2:D4', { mode });
      });
      const range = { startRow: 0, endRow: 5, startColumn: 0, endColumn: 4 };
      const read = async (id: string) =>
        (
          await book.readCells(id, range, {
            fields: ['input', 'value', 'style', 'numberFormat', 'note', 'link'],
          })
        ).cells;
      results.push({ mode, direct: await read(direct.id), staged: await read(staged.id) });
    }
    return results;
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  for (const item of result) expect(item.staged, item.mode).toEqual(item.direct);
});
