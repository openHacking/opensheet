import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('oversized bulk command fails with a budget error without changing the workbook', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([['original']]);
      const revision = book.revision;
      let code = '';
      let message = '';
      try {
        await sheet.range('A1:A100000').setValues(Array.from({ length: 100000 }, () => ['new']));
      } catch (error) {
        code = (error as { code?: string }).code ?? '';
        message = String(error);
      }
      const after = {
        revision: book.revision,
        first: await sheet.range('A1').getValues(),
        last: await sheet.range('A100000').getValues(),
      };
      const undo = await book.undo();
      return { code, message, revision, after, undo, undone: await sheet.range('A1').getValues() };
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Bulk', rows: 100000, columns: 26 }] })),
  );
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.message).not.toContain('call stack');
  expect(result.after).toEqual({
    revision: result.revision,
    first: [['original']],
    last: [[null]],
  });
  expect(result.undo).toBe(true);
  expect(result.undone).toEqual([[null]]);
});
