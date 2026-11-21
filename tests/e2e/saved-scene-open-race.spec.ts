import { expect, test } from '@playwright/test';

for (const race of ['reference', 'reference-cleanup', 'edit', 'draft', 'read-error'] as const) {
  test(`opening a saved scene preserves the current workbook after ${race} changes`, async ({
    page,
    context,
  }) => {
    page.on('dialog', (dialog) => void dialog.accept());
    await page.goto('/#budget');
    await expect(page.locator('#scene-select')).toBeEnabled();
    await page.locator('#scene-select').selectOption('sales');
    await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
    const savedId = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[1111]]);
      return book.id;
    });
    await page.locator('#scene-select').selectOption('budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const other = await context.newPage();
    await other.goto('/#budget');
    await expect(other.locator('#scene-select')).toBeEnabled();
    const otherId = await other.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        snapshot = await book.toJSON();
      const Workbook = book.constructor as new (file: typeof snapshot) => typeof book;
      const replacement = new Workbook({ ...snapshot, workbookId: `other_${crypto.randomUUID()}` });
      await replacement.ready();
      await replacement
        .getSheets()[0]
        .range('A1')
        .setValues([[8642]]);
      await replacement.close();
      return replacement.id;
    });
    const original = await page.evaluate((id) => {
      const book = window.opensheet.getWorkbook(),
        prototype = Object.getPrototypeOf(book),
        ready = prototype.ready;
      prototype.ready = async function () {
        const result = await ready.call(this);
        if (this !== book && this.id === id) {
          prototype.ready = ready;
          (window as any).candidateScene = this;
          await new Promise<void>((resolve) => ((window as any).resumeSavedScene = resolve));
        }
        return result;
      };
      return { id: book.id, revision: book.revision };
    }, savedId);
    await page.locator('#scene-select').selectOption('sales');
    await expect.poll(() => page.evaluate(() => !!(window as any).resumeSavedScene)).toBe(true);
    if (race === 'reference' || race === 'reference-cleanup')
      await other.evaluate((id) => localStorage.setItem('opensheet.scene.sales', id), otherId);
    if (race === 'reference-cleanup')
      await page.evaluate(() => {
        const candidate = (window as any).candidateScene;
        const close = candidate.close.bind(candidate);
        candidate.close = async () => {
          await close();
          throw new Error('Injected cleanup failure');
        };
      });
    if (race === 'edit')
      await page.evaluate(async () =>
        window.opensheet
          .getWorkbook()
          .getSheets()[0]
          .range('A1')
          .setValues([[2468]]),
      );
    if (race === 'draft') {
      await page.evaluate(() => {
        const book = window.opensheet.getWorkbook();
        window.opensheet.selection.set({
          sheetId: book.getSheets()[0].id,
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 1,
        });
      });
      await page.getByRole('grid', { exact: true }).focus();
      await page.getByRole('grid', { exact: true }).press('Enter');
      await page.getByRole('textbox', { name: 'Edit cell', exact: true }).fill('2468');
      await expect.poll(() => page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
    }
    if (race === 'read-error')
      await page.evaluate(() => {
        const originalGet = Storage.prototype.getItem;
        Storage.prototype.getItem = function (key) {
          if (this === localStorage && key === 'opensheet.scene.sales') {
            Storage.prototype.getItem = originalGet;
            throw new DOMException('Reference read denied', 'SecurityError');
          }
          return originalGet.call(this, key);
        };
      });
    await page.evaluate(() => (window as any).resumeSavedScene());
    await expect(page.locator('#notice')).toContainText(
      race === 'reference' || race === 'reference-cleanup'
        ? 'Saved scene changed'
        : race === 'edit' || race === 'draft'
          ? 'Workbook changed'
          : 'Reference read denied',
      { timeout: 5000 },
    );
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const state = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        id: book.id,
        revision: book.revision,
        value: await book.getSheets()[0].range('A1').getValues(),
        reference: localStorage.getItem('opensheet.scene.sales'),
      };
    });
    expect(state.id).toBe(original.id);
    expect(state.revision).toBe(original.revision + (race === 'edit' ? 1 : 0));
    expect(state.reference).toBe(
      race === 'reference' || race === 'reference-cleanup' ? otherId : savedId,
    );
    if (race === 'edit') expect(state.value).toEqual([[2468]]);
    if (race === 'draft') {
      expect(await page.evaluate(() => window.opensheet.hasDraft)).toBe(true);
      const editor = page.getByRole('textbox', { name: 'Edit cell', exact: true });
      await expect(editor).toHaveValue('2468');
      await editor.press('Escape');
    }
    const candidate = await other.evaluate(
      async ({ modulePath, id }) => {
        const { openWorkbook } = await import(/* @vite-ignore */ modulePath);
        const book = await openWorkbook(id);
        try {
          return {
            readOnly: book.isReadOnly,
            value: await book.getSheets()[0].range('A1').getValues(),
          };
        } finally {
          await book.close();
        }
      },
      { modulePath: `/@fs${process.cwd()}/packages/core/src/index.ts`, id: savedId },
    );
    expect(candidate).toEqual({ readOnly: false, value: [[1111]] });
    await page.locator('#scene-select').selectOption('sales');
    await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
    const opened = await page.evaluate(async () => ({
      id: window.opensheet.getWorkbook().id,
      value: await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    }));
    expect(opened).toEqual({
      id: race === 'reference' || race === 'reference-cleanup' ? otherId : savedId,
      value: [[race === 'reference' || race === 'reference-cleanup' ? 8642 : 1111]],
    });
    await other.close();
  });
}

test('switching to another scene with the same saved workbook keeps its writer', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const original = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[731]]);
    localStorage.setItem('opensheet.scene.sales', book.id);
    (window as any).originalSceneBook = book;
    return book.id;
  });
  page.on('dialog', (dialog) => void dialog.accept());
  await page.locator('#scene-select').selectOption('sales');
  await expect(page.locator('#scene-title')).toHaveText('Sales Dashboard');
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return {
        id: book.id,
        same: book === (window as any).originalSceneBook,
        readOnly: book.isReadOnly,
        value: await book.getSheets()[0].range('A1').getValues(),
      };
    }),
  ).toEqual({ id: original, same: true, readOnly: false, value: [[731]] });
});
