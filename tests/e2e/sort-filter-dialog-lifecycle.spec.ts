import { expect, test } from '@playwright/test';

for (const [action, other, button, failure] of [
  ['sort-custom', 'filter-custom', 'Apply sort', false],
  ['sort-custom', 'filter-custom', 'Apply sort', true],
  ['filter-custom', 'sort-custom', 'Apply filter', false],
  ['filter-custom', 'sort-custom', 'Apply filter', true],
  ['filter-custom', 'sort-custom', 'Clear all filters', false],
] as const) {
  test(`${action} ${button} ${failure ? 'failure' : 'completion'} does not affect a later dialog`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page.evaluate(
      async ({ fail, clear }) => {
        const app = window.opensheet;
        const book = app.getWorkbook();
        const sheet = book.getSheets()[0];
        await sheet.range('A1:A3').setValues([[3], [1], [2]]);
        if (clear) await sheet.setFilter({ column: 0, query: '1' });
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 3,
          startColumn: 0,
          endColumn: 1,
        });
        const execute = book.execute.bind(book);
        let release!: () => void;
        let committed!: () => void;
        const wait = new Promise<void>((resolve) => (release = resolve));
        const written = new Promise<void>((resolve) => (committed = resolve));
        (window as any).sortFilterGate = { release, written };
        book.execute = async (...args: Parameters<typeof book.execute>) => {
          const result = await execute(...args);
          committed();
          await wait;
          if (fail) throw new Error('Delayed failure after commit');
          return result;
        };
      },
      { fail: failure, clear: button === 'Clear all filters' },
    );
    await page.getByLabel('Sheet actions').selectOption(action);
    await page.getByRole('button', { name: button }).click();
    await page.evaluate(() => (window as any).sortFilterGate.written);
    const committed = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      return {
        values: await sheet.range('A1:A3').getValues(),
        filter: book.sheetData(sheet.id).filter,
      };
    });
    if (action === 'sort-custom') expect(committed.values).toEqual([[1], [2], [3]]);
    else if (button === 'Clear all filters') expect(committed.filter).toBeNull();
    else expect(committed.filter).not.toBeNull();
    await page.locator('#dialog-close').click();
    await page.getByLabel('Sheet actions').selectOption(other);
    await expect(page.getByRole('dialog')).toBeVisible();
    await page.evaluate(async () => {
      const notice = document.querySelector<HTMLElement>('#notice')!;
      notice.hidden = true;
      notice.textContent = '';
      (window as any).sortFilterGate.release();
      await new Promise((resolve) => setTimeout(resolve, 0));
    });
    await expect(page.getByRole('dialog')).toBeVisible();
    await expect(page.locator('#notice')).toBeHidden();
  });
}
