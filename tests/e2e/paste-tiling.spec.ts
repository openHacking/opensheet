import { expect, test } from '@playwright/test';
import { readForPoll } from './read-for-poll.js';

const moduleUrl = `/@fs${process.cwd()}/packages/renderer/src/clipboard.ts`;

test('tiles structured clipboard data with per-cell formula relocation and one undo', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { pasteClipboard } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:F4').clear({ all: true });
    await sheet.range('A1').setValues([[4]]);
    await sheet.range('A1').setNote('source');
    await sheet.range('A1').setStyle({ bold: true });
    await sheet.range('B1').setFormulas([['A1+$A$1']]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 });
    const source = await app.getGrid()!.copyClipboard();
    const selection = { sheetId: sheet.id, startRow: 2, endRow: 4, startColumn: 2, endColumn: 6 };
    const pasted = await pasteClipboard(book, selection, source);
    const values = await sheet.range('C3:F4').getValues();
    const formulas = await sheet.range('C3:F4').getFormulas();
    const notes = await sheet.range('C3:F4').getNotes();
    const styles = await Promise.all(
      [2, 4].map(async (column) =>
        book.getStyle((await book.getCell(sheet.id, 2, column))?.styleId),
      ),
    );
    await book.undo();
    const undone = await sheet.range('C3:F4').getValues();
    await book.redo();
    await app.open(book.id);
    return {
      pasted,
      selection,
      values,
      formulas,
      notes,
      styles,
      undone,
      reopened: await app.getWorkbook().getSheets()[0].range('C3:F4').getValues(),
    };
  }, moduleUrl);
  expect(result.pasted).toEqual(result.selection);
  expect(result.values).toEqual([
    [4, 8, 4, 8],
    [4, 8, 4, 8],
  ]);
  expect(result.formulas).toEqual([
    [null, '(C3+$A$1)', null, '(E3+$A$1)'],
    [null, '(C4+$A$1)', null, '(E4+$A$1)'],
  ]);
  expect(result.notes).toEqual([
    ['source', null, 'source', null],
    ['source', null, 'source', null],
  ]);
  expect(result.styles).toEqual([{ bold: true }, { bold: true }]);
  expect(result.undone).toEqual([
    [null, null, null, null],
    [null, null, null, null],
  ]);
  expect(result.reopened).toEqual(result.values);
});

test('tiles text and format-only HTML and rejects oversized expansion atomically', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { pasteClipboard } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:D4').clear({ all: true });
    const selection = { sheetId: sheet.id, startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 };
    await pasteClipboard(book, selection, { text: '1\t2\n3\t4' });
    const values = await sheet.range('A1:D4').getValues();
    await pasteClipboard(
      book,
      selection,
      { text: '', html: '<table><tr><td style="font-weight:bold">x</td></tr></table>' },
      { mode: 'formats' },
    );
    const styles = await Promise.all(
      [
        [0, 0],
        [3, 3],
      ].map(async ([row, col]) => book.getStyle((await book.getCell(sheet.id, row, col))?.styleId)),
    );
    const revision = book.revision;
    const failure = await pasteClipboard(
      book,
      { sheetId: sheet.id, startRow: 0, endRow: 100000, startColumn: 0, endColumn: 3 },
      { text: 'x' },
    ).then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    return {
      values,
      styles,
      failure,
      revision,
      after: book.revision,
      unchanged: await sheet.range('A1:D4').getValues(),
    };
  }, moduleUrl);
  expect(result.values).toEqual([
    [1, 2, 1, 2],
    [3, 4, 3, 4],
    [1, 2, 1, 2],
    [3, 4, 3, 4],
  ]);
  expect(result.styles).toEqual([{ bold: true }, { bold: true }]);
  expect(result.failure).toBe('LIMIT_EXCEEDED');
  expect(result.after).toBe(result.revision);
  expect(result.unchanged).toEqual(result.values);
});

test('uses source size for mismatched selection, includes hidden rows and rejects merged targets', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { pasteClipboard } = await import(/* @vite-ignore */ url);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:D4').clear({ all: true });
    await sheet.setRowHidden(1, true);
    await sheet.setColumnHidden(1, true);
    const selection = { sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 };
    const pasted = await pasteClipboard(book, selection, { text: '1\t2\n3\t4' });
    const values = await sheet.range('A1:C3').getValues();
    await sheet.range('B2:C2').merge();
    const revision = book.revision;
    const failure = await pasteClipboard(
      book,
      { ...selection, endRow: 2, endColumn: 2 },
      { text: '7' },
    ).then(
      () => 'success',
      (error: { code: string }) => error.code,
    );
    return {
      pasted,
      values,
      revision,
      after: book.revision,
      failure,
      final: await sheet.range('A1:C3').getValues(),
    };
  }, moduleUrl);
  expect(result.pasted).toMatchObject({ endRow: 2, endColumn: 2 });
  expect(result.values).toEqual([
    [1, 2, null],
    [3, 4, null],
    [null, null, null],
  ]);
  expect(result.failure).toBe('INVALID_RANGE');
  expect(result.after).toBe(result.revision);
  expect(result.final).toEqual(result.values);
});

test('grid paste fills the selected rectangle from one text cell', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:C3').clear({ all: true });
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 3, startColumn: 0, endColumn: 3 });
    const data = new DataTransfer();
    data.setData('text/plain', '42');
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: data });
    app.getGrid()!.element.dispatchEvent(event);
  });
  await expect
    .poll(() => readForPoll(page, 'A1:C3'))
    .toEqual([
      [42, 42, 42],
      [42, 42, 42],
      [42, 42, 42],
    ]);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 0,
    endRow: 3,
    startColumn: 0,
    endColumn: 3,
  });
});
