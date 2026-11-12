import { expect, test } from '@playwright/test';

test('modified shortcuts do not trigger unrelated spreadsheet mutations', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 8, columns: 8 }] }));
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:A2').setValues([[1], [2]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 });
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  const result = await page.evaluate(() => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const revision = book.revision;
    const originalExecute = book.execute.bind(book);
    const originalUndo = book.undo.bind(book);
    let writes = 0;
    let undos = 0;
    (book as any).execute = (...args: unknown[]) => {
      writes++;
      return (originalExecute as any)(...args);
    };
    (book as any).undo = (...args: unknown[]) => {
      undos++;
      return (originalUndo as any)(...args);
    };
    const element = app.getGrid()!.element;
    for (const key of ['d', 'z', 'y', 'a', 'Delete', 'Backspace'])
      element.dispatchEvent(
        new KeyboardEvent('keydown', { key, ctrlKey: true, altKey: true, bubbles: true }),
      );
    for (const key of ['Delete', 'Backspace'])
      element.dispatchEvent(new KeyboardEvent('keydown', { key, shiftKey: true, bubbles: true }));
    return { revision, current: book.revision, selection: app.selection.get(), writes, undos };
  });
  expect(result).toMatchObject({
    current: result.revision,
    selection: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
    writes: 0,
    undos: 0,
  });
  const readValues = () =>
    page.evaluate(async () => {
      try {
        return await window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues();
      } catch (error) {
        if ((error as { code?: string }).code === 'REVISION_CONFLICT') return undefined;
        throw error;
      }
    });
  await page.keyboard.press('Control+d');
  await expect.poll(readValues).toEqual([[1], [1]]);
  await page.keyboard.press('Control+z');
  await expect.poll(readValues).toEqual([[1], [2]]);
  await page.keyboard.press('Delete');
  await expect.poll(readValues).toEqual([[null], [null]]);
});

test('unhandled modified navigation is not mistaken for plain grid movement or editing', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const outcome = await page.evaluate(() => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const before = app.selection.get();
    const revision = book.revision;
    const grid = app.getGrid()!.element;
    const cases = [
      { key: 'ArrowUp', altKey: true },
      { key: 'Tab', ctrlKey: true },
      { key: 'Home', altKey: true },
      { key: 'Enter', altKey: true },
      { key: 'F2', ctrlKey: true },
    ];
    const outcomes = cases.map((init) => {
      const event = new KeyboardEvent('keydown', { ...init, bubbles: true, cancelable: true });
      grid.dispatchEvent(event);
      return {
        prevented: event.defaultPrevented,
        selection: app.selection.get(),
        editing: !grid.querySelector<HTMLTextAreaElement>('.os-cell-editor')!.hidden,
      };
    });
    return {
      before,
      after: app.selection.get(),
      revision,
      current: book.revision,
      outcomes,
      editing: !grid.querySelector<HTMLTextAreaElement>('.os-cell-editor')!.hidden,
    };
  });
  expect(outcome.after).toEqual(outcome.before);
  expect(outcome.current).toBe(outcome.revision);
  expect(outcome.outcomes).toEqual(
    Array.from({ length: 5 }, () => ({
      prevented: false,
      selection: outcome.before,
      editing: false,
    })),
  );
  expect(outcome.editing).toBe(false);
});
