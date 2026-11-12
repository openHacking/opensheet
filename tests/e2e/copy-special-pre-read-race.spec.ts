import { expect, test } from '@playwright/test';

for (const mode of ['values', 'formulas', 'formats'] as const) {
  test(`${mode} copy rejects a source read from a newer revision`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(
      async ({ url, mode }) => {
        const { createWorkbookFile } = await import(/* @vite-ignore */ url);
        const app = window.opensheet;
        await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
        let book = app.getWorkbook();
        let sheet = book.getSheets()[0];
        await sheet.range('A1').setFormulas([['1+2']]);
        await sheet.range('A1').setStyle({ bold: true });
        await sheet.range('B1').setValues([[99]]);
        const read = book.readCells.bind(book);
        let release!: () => void;
        const gate = new Promise<void>((resolve) => {
          release = resolve;
        });
        book.readCells = async (...args) => {
          await gate;
          return read(...args);
        };
        const pending = sheet
          .range('A1')
          .copyTo('B1', { mode })
          .then(
            () => 'success',
            (error) => error.code,
          );
        await sheet.insertRows(0);
        const revision = book.revision;
        release();
        const code = await pending;
        book.readCells = read;
        const after = { revision: book.revision, values: await sheet.range('A1:B2').getValues() };
        await app.open(book.id);
        book = app.getWorkbook();
        sheet = book.getSheets()[0];
        const reopened = await sheet.range('A1:B2').getValues();
        await book.undo();
        const undone = await sheet.range('A1:B1').getValues();
        await book.redo();
        return {
          code,
          revision,
          after,
          reopened,
          undone,
          redone: await sheet.range('A1:B2').getValues(),
        };
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, mode },
    );
    expect(result.code).toBe('REVISION_CONFLICT');
    expect(result.after).toEqual({
      revision: result.revision,
      values: [
        [null, null],
        [3, 99],
      ],
    });
    expect(result.reopened).toEqual(result.after.values);
    expect(result.undone).toEqual([[3, 99]]);
    expect(result.redone).toEqual(result.after.values);
  });
}
