import { expect, test } from '@playwright/test';

test('large original-input preconditions reject without publishing or changing history', async ({
  page,
}) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Sheet', rows: 600, columns: 1 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[42]]);
      const revision = book.revision;
      const bytes = (await book.storageStats()).bytes;
      const data = book.sheetData(sheet.id);
      const text = 'x'.repeat(32_768);
      const command = {
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: Array.from({ length: 512 }, (_, row) => ({
            row,
            column: 0,
            input: { type: 'number', value: 1 },
            expected: {
              rowId: data.rowOrder[row],
              columnId: data.columnOrder[0],
              input: { type: 'string', value: text },
            },
          })),
        },
      };
      const code = await book.execute(command).then(
        () => 'success',
        (error: { code?: string }) => error.code,
      );
      const unchanged = {
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        value: await sheet.range('A1').getValues(),
      };
      await book.undo();
      const undone = await sheet.range('A1').getValues();
      await book.redo();
      return {
        code,
        revision,
        bytes,
        unchanged,
        undone,
        redone: await sheet.range('A1').getValues(),
      };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.unchanged).toEqual({
    revision: result.revision,
    bytes: result.bytes,
    value: [[42]],
  });
  expect(result.undone).toEqual([[null]]);
  expect(result.redone).toEqual([[42]]);
});
