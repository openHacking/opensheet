import { expect, test } from '@playwright/test';

test('oversized relocated formula rejects the entire tile across reopen and history', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[7, 8]]);
    const revision = book.revision;
    let code = '';
    try {
      await book.execute({
        type: 'core.cells.pasteTile',
        payload: {
          sheetId: sheet.id,
          range: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
          rows: 1,
          columns: 2,
          cells: [
            { input: { type: 'number', value: 42 } },
            {
              input: { type: 'formula', expression: `A1&"${'x'.repeat(32763)}"` },
              source: { row: 0, column: 0 },
            },
          ],
        },
      });
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    const after = { revision: book.revision, values: await sheet.range('A1:B1').getValues() };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1:B1').getValues();
    await book.undo();
    await book.redo();
    return { code, revision, after, reopened, redone: await sheet.range('A1:B1').getValues() };
  });
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.after.revision).toBe(result.revision);
  expect(result.after.values).toEqual([[7, 8]]);
  expect(result.reopened).toEqual([[7, 8]]);
  expect(result.redone).toEqual([[7, 8]]);
});
