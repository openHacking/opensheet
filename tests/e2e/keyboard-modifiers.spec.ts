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

test('modified navigation leaves an in-cell draft and browser shortcuts alone', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const grid = page.getByRole('grid', { exact: true });
  await grid.focus();
  await page.keyboard.press('F2');
  const editor = page.locator('.os-cell-editor');
  await expect(editor).toBeVisible();
  await editor.fill('draft');
  const result = await page.evaluate(() => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const input = document.querySelector<HTMLTextAreaElement>('.os-cell-editor')!;
    const before = { selection: app.selection.get(), revision: book.revision };
    const cases = [
      { key: 'Tab', ctrlKey: true },
      { key: 'Tab', metaKey: true, shiftKey: true },
      { key: 'Tab', altKey: true },
      { key: 'Enter', ctrlKey: true },
      { key: 'Enter', metaKey: true },
      { key: 'Enter', altKey: true, ctrlKey: true },
    ];
    const outcomes = cases.map((init) => {
      const event = new KeyboardEvent('keydown', { ...init, bubbles: true, cancelable: true });
      input.dispatchEvent(event);
      return {
        prevented: event.defaultPrevented,
        editing: !input.hidden,
        draft: input.value,
        selection: app.selection.get(),
        revision: book.revision,
      };
    });
    return { before, outcomes };
  });
  for (const outcome of result.outcomes)
    expect(outcome).toEqual({ prevented: false, editing: true, draft: 'draft', ...result.before });
  await expect(editor).toHaveValue('draft');
  await editor.press('Alt+Enter');
  await expect(editor).toHaveValue('draft\n');
  await editor.press('Enter');
  await expect(editor).toBeHidden();
  await expect
    .poll(() =>
      page.evaluate(async () => {
        try {
          return await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues();
        } catch (error) {
          if ((error as { code?: string }).code === 'REVISION_CONFLICT') return undefined;
          throw error;
        }
      }),
    )
    .toEqual([['draft\n']]);
});

test('modified Enter does not submit a formula-bar draft', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const input = page.getByRole('combobox', { name: 'Formula bar', exact: true });
  await input.fill('=1+2');
  const result = await input.evaluate(async (element: HTMLInputElement) => {
    const book = window.opensheet.getWorkbook();
    const revision = book.revision;
    const original = book.execute.bind(book);
    let writes = 0;
    (book as any).execute = (...args: unknown[]) => {
      writes++;
      return (original as any)(...args);
    };
    const outcomes = [{ ctrlKey: true }, { metaKey: true }, { altKey: true }].map((modifiers) => {
      const event = new KeyboardEvent('keydown', {
        key: 'Enter',
        ...modifiers,
        bubbles: true,
        cancelable: true,
      });
      element.dispatchEvent(event);
      return { prevented: event.defaultPrevented, value: element.value, revision: book.revision };
    });
    await new Promise((resolve) => setTimeout(resolve, 100));
    (book as any).execute = original;
    return { outcomes, revision, writes };
  });
  expect(result.writes).toBe(0);
  expect(result.outcomes).toEqual(
    Array.from({ length: 3 }, () => ({
      prevented: false,
      value: '=1+2',
      revision: result.revision,
    })),
  );
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(result.revision + 1);
});

test('modified Enter does not submit the address box', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const input = page.locator('.os-namebox');
  await input.fill('B2');
  const before = await page.evaluate(() => window.opensheet.selection.get());
  const prevented = await input.evaluate((element: HTMLInputElement) =>
    [{ ctrlKey: true }, { metaKey: true }, { altKey: true }].map((modifiers) => {
      const event = new KeyboardEvent('keydown', {
        key: 'Enter',
        ...modifiers,
        bubbles: true,
        cancelable: true,
      });
      element.dispatchEvent(event);
      return event.defaultPrevented;
    }),
  );
  await page.waitForTimeout(100);
  expect(prevented).toEqual([false, false, false]);
  expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(before);
  await expect(input).toHaveValue('B2');
  await input.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.selection.get()))
    .toMatchObject({ startRow: 1, startColumn: 1 });
});
