import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const original = book.getCell.bind(book);
    (window as any).originalEditorRead = original;
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    (window as any).releaseEditorRead = release;
    (window as any).editorReadStarted = false;
    book.getCell = async (...args) => {
      const cell = await original(...args);
      (window as any).editorReadStarted = true;
      await gate;
      return cell;
    };
  });
});

test('late cell reads cannot reopen editing after selection changes', async ({ page }) => {
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.evaluate(() => {
    const s = window.opensheet.selection.get()!;
    window.opensheet.selection.set({ ...s, startRow: 1, endRow: 2, startColumn: 1, endColumn: 2 });
    (window as any).releaseEditorRead();
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.getByRole('textbox', { name: 'Edit cell' })).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 1,
    startColumn: 1,
  });
});

test('late reads cannot replace a newer typed draft', async ({ page }) => {
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.keyboard.type('x');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await expect(editor).toHaveValue('x');
  await page.evaluate(() => (window as any).releaseEditorRead());
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(editor).toHaveValue('x');
  await editor.press('Escape');
});

test('returning to the original cell does not revive a cancelled read', async ({ page }) => {
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.evaluate(() => {
    const original = window.opensheet.selection.get()!;
    window.opensheet.selection.set({
      ...original,
      startRow: 1,
      endRow: 2,
      startColumn: 1,
      endColumn: 2,
    });
    window.opensheet.selection.set(original);
    (window as any).releaseEditorRead();
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.getByRole('textbox', { name: 'Edit cell' })).toBeHidden();
  expect(await page.evaluate(() => window.opensheet.selection.get())).toMatchObject({
    startRow: 0,
    startColumn: 0,
  });
});

test('Escape cancels a pending read', async ({ page }) => {
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.keyboard.press('Escape');
  await page.evaluate(() => {
    (window as any).releaseEditorRead();
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.getByRole('textbox', { name: 'Edit cell' })).toBeHidden();
});

test('read-only mode cancels a pending read', async ({ page }) => {
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.evaluate(() => {
    window.opensheet.setMode('read');
    (window as any).releaseEditorRead();
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.getByRole('textbox', { name: 'Edit cell' })).toBeHidden();
});

test('focus changes cancel pending reads', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  const address = page.getByRole('textbox', { name: 'Cell address' });
  await address.focus();
  await page.evaluate(() => (window as any).releaseEditorRead());
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(address).toBeFocused();
  await expect(page.getByRole('textbox', { name: 'Edit cell' })).toBeHidden();
  expect(errors).toEqual([]);
});

test('disposed grids ignore pending read completion', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.evaluate(() => {
    window.opensheet.dispose();
    (window as any).releaseEditorRead();
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.locator('.os-cell-editor')).toHaveCount(0);
  expect(errors).toEqual([]);
});

test('sheet switching cancels a pending read without returning to the old sheet', async ({
  page,
}) => {
  await page.keyboard.press('F2');
  await page.waitForFunction(() => (window as any).editorReadStarted);
  await page.getByRole('tab', { name: 'Notes', exact: true }).click();
  await page.evaluate(() => (window as any).releaseEditorRead());
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.getByRole('tab', { name: 'Notes', exact: true })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(page.getByRole('textbox', { name: 'Edit cell' })).toBeHidden();
});

test('active read failures are reported but cancelled failures are silent', async ({ page }) => {
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    let reject!: (error: Error) => void;
    const pending = new Promise<never>((_, fail) => {
      reject = fail;
    });
    (window as any).rejectEditorRead = reject;
    book.getCell = () => pending;
  });
  await page.keyboard.press('F2');
  await page.keyboard.press('Escape');
  await page.evaluate(() => (window as any).rejectEditorRead(new Error('cancelled read')));
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(page.locator('.os-toast')).toBeHidden();
  await page.evaluate(() => {
    window.opensheet.getWorkbook().getCell = async () => {
      throw new Error('active read');
    };
  });
  await page.keyboard.press('F2');
  await expect(page.locator('.os-toast')).toHaveText('active read');
});

test('only the latest of two out-of-order cell reads may open the editor', async ({ page }) => {
  await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const current: typeof book.getCell = (window as any).originalEditorRead;
    const releases: Array<() => void> = [];
    const gates = [0, 1].map(() => new Promise<void>((resolve) => releases.push(resolve)));
    (window as any).readReleases = releases;
    (window as any).readCount = 0;
    book.getCell = async (...args) => {
      const index = (window as any).readCount++;
      const cell = await current(...args);
      await gates[index];
      return { ...cell, input: { type: 'string', value: index === 0 ? 'old' : 'new' } };
    };
    window.opensheet.getGrid()!.startEditing();
    window.opensheet.getGrid()!.startEditing();
  });
  await page.waitForFunction(() => (window as any).readCount === 2);
  await page.evaluate(() => (window as any).readReleases[1]());
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await expect(editor).toHaveValue('new');
  await page.evaluate(() => (window as any).readReleases[0]());
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(editor).toHaveValue('new');
  await editor.press('Escape');
});

test('a redundant start does not replace an active draft', async ({ page }) => {
  await page.keyboard.type('x');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await expect(editor).toHaveValue('x');
  await page.evaluate(() => {
    window.opensheet.getGrid()!.startEditing();
    (window as any).releaseEditorRead();
  });
  await page.evaluate(
    () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve))),
  );
  await expect(editor).toHaveValue('x');
  expect(await page.evaluate(() => (window as any).editorReadStarted)).toBe(false);
  await editor.press('Escape');
});
