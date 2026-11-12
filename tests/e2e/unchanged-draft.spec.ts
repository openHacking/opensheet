import { expect, test } from '@playwright/test';

for (const surface of ['cell', 'formula'] as const) {
  test(`unchanged ${surface} draft preserves revision and redo`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const revision = await page.evaluate(async (url) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[10]]);
      await sheet.range('A1').setValues([[20]]);
      await book.undo();
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      await app.getGrid()!.ready();
      return book.revision;
    }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
    if (surface === 'cell') {
      const grid = page.getByRole('grid', { exact: true });
      await grid.focus();
      await grid.press('F2');
      const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
      await expect(editor).toHaveValue('10');
      await editor.press('Enter');
    } else {
      const editor = page.getByLabel('Formula bar', { exact: true });
      await expect(editor).toHaveValue('10');
      await editor.focus();
      await editor.press('Enter');
    }
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending))
      .toBe(0);
    const result = await page.evaluate(async () => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const after = book.revision;
      const redone = await book.redo();
      return { after, redone, value: await book.getSheets()[0].range('A1').getValues() };
    });
    expect(result).toEqual({ after: revision, redone: true, value: [[20]] });
  });
}

for (const surface of ['cell', 'formula'] as const) {
  test(`unchanged ${surface} draft preserves literal formula-like text`, async ({ page }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const revision = await page.evaluate(async (url) => {
      const { createWorkbookFile } = await import(/* @vite-ignore */ url);
      const app = window.opensheet;
      await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
      const book = app.getWorkbook();
      const sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([['=1+1']]);
      app.selection.set({
        sheetId: sheet.id,
        startRow: 0,
        endRow: 1,
        startColumn: 0,
        endColumn: 1,
      });
      await app.getGrid()!.ready();
      return book.revision;
    }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
    const editor =
      surface === 'cell'
        ? page.getByRole('textbox', { name: 'Edit cell', exact: true })
        : page.getByLabel('Formula bar', { exact: true });
    if (surface === 'cell') {
      const grid = page.getByRole('grid', { exact: true });
      await grid.focus();
      await grid.press('F2');
    } else await editor.focus();
    await expect(editor).toHaveValue('=1+1');
    await editor.press('Enter');
    await expect
      .poll(() => page.evaluate(() => window.opensheet.getWorkbook().persistenceState.pending))
      .toBe(0);
    const result = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      return { revision: book.revision, input: (await book.getCell(sheet.id, 0, 0))?.input };
    });
    expect(result).toEqual({ revision, input: { type: 'string', value: '=1+1' } });
  });
}
