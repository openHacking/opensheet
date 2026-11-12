import { expect, test } from '@playwright/test';

for (const operation of ['fill', 'sort'] as const) {
  test(`${operation} rejects an oversized relocated formula without publication`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(
      async ({ url, operation }) => {
        const { createWorkbookFile } = await import(/* @vite-ignore */ url);
        const app = window.opensheet;
        await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
        let book = app.getWorkbook();
        let sheet = book.getSheets()[0];
        const expression = `A1&"${'x'.repeat(32763)}"`;
        await book.execute({
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [
              { row: 0, column: 0, input: { type: 'number', value: 2 } },
              { row: 1, column: 0, input: { type: 'number', value: 1 } },
              { row: 0, column: 1, input: { type: 'formula', expression } },
            ],
          },
        });
        const revision = book.revision;
        const bytes = (await book.storageStats()).bytes;
        const failure = await (
          operation === 'fill' ? sheet.range('A1:B2').fillDown() : sheet.sort('A1:B2', 0)
        ).then(
          () => 'success',
          (error) => error.code,
        );
        const current = {
          revision: book.revision,
          bytes: (await book.storageStats()).bytes,
          inputs: await sheet.range('A1:B2').getInputs(),
        };
        await app.open(book.id);
        book = app.getWorkbook();
        sheet = book.getSheets()[0];
        const reopened = await sheet.range('A1:B2').getInputs();
        await book.undo();
        await book.redo();
        return {
          failure,
          revision,
          bytes,
          current,
          reopened,
          redone: await sheet.range('A1:B2').getInputs(),
          expression,
        };
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, operation },
    );
    expect(result.failure).toBe('LIMIT_EXCEEDED');
    expect(result.current.revision).toBe(result.revision);
    expect(result.current.bytes).toBe(result.bytes);
    expect(result.current.inputs[0][1]).toEqual({ type: 'formula', expression: result.expression });
    expect(result.reopened).toEqual(result.current.inputs);
    expect(result.redone).toEqual(result.current.inputs);
  });
}
