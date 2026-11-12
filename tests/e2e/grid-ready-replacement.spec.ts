import { expect, test } from '@playwright/test';

for (const mode of ['reopen', 'load'] as const)
  test(`${mode} settles old grid readiness and ignores a late prefetch error`, async ({ page }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (replacement) => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([['before reopen']]);
      const snapshot = await book.toJSON();
      const grid = app.getGrid()!;
      await grid.ready();
      while ((grid as any).fetching) await (grid as any).fetching;
      let entered!: () => void;
      let rejectRead!: (error: Error) => void;
      const started = new Promise<void>((resolve) => (entered = resolve));
      const stalled = new Promise<void>((_resolve, reject) => (rejectRead = reject));
      const errors: string[] = [];
      (grid as any).options.onError = (error: Error) => errors.push(error.message);
      book.prefetch = () => {
        entered();
        return stalled;
      };
      const oldReady = grid.ready();
      await started;
      if (replacement === 'reopen') await app.open(book.id);
      else await app.load(snapshot);
      const reopened = app.getWorkbook();
      const settled = await Promise.race([
        oldReady.then(() => true),
        new Promise<false>((resolve) => setTimeout(() => resolve(false), 500)),
      ]);
      rejectRead(new Error('late old-grid failure'));
      await new Promise((resolve) => setTimeout(resolve, 0));
      const sheet = reopened.getSheets()[0];
      const restored = await sheet.range('A1').getValues();
      await sheet.range('A1').setValues([['after reopen']]);
      await app.ready();
      return {
        settled,
        errors,
        replaced: app.getGrid() !== grid && app.getWorkbook() === reopened,
        restored,
        edited: await sheet.range('A1').getValues(),
      };
    }, mode);
    expect(result).toEqual({
      settled: true,
      errors: [],
      replaced: true,
      restored: [['before reopen']],
      edited: [['after reopen']],
    });
  });
