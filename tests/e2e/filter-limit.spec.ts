import { expect, test } from '@playwright/test';

test('legacy and typed filters reject calculation exhaustion atomically', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const initial = app.getWorkbook();
    const sheet = await initial.addSheet('Filter limit', { rows: 132, columns: 2 });
    await sheet.range('B1').setValues([[1]]);
    await sheet.range('A132').setValues([[1]]);
    await sheet
      .range('A2:A131')
      .setFormulas(Array.from({ length: 130 }, (_, i) => [`A${i + 3}+1`]));
    await app.open(initial.id);
    const book = app.getWorkbook();
    const target = book.getSheetByName('Filter limit')!;
    const before = {
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      canUndo: book.canUndo,
      canRedo: book.canRedo,
    };
    const failures: string[] = [];
    const states: unknown[] = [];
    for (const filter of [
      { column: 0, query: '1' },
      {
        mode: 'all' as const,
        headerRows: 1,
        criteria: [{ column: 0, operator: 'gt' as const, value: 0 }],
      },
    ]) {
      try {
        await book.transaction({}, async (tx) => {
          const staged = tx.getSheetByName('Filter limit')!;
          await staged.range('B1').setValues([[999]]);
          await staged.setFilter(filter);
        });
        failures.push('unexpected success');
      } catch (error) {
        failures.push((error as { code?: string }).code ?? 'unknown');
      }
      states.push({
        revision: book.revision,
        bytes: (await book.storageStats()).bytes,
        canUndo: book.canUndo,
        canRedo: book.canRedo,
        value: await target.range('B1').getValues(),
        filter: book.metadata.sheets.find((item) => item.id === target.id)?.filter,
        hidden: book.filteredRows(target.id),
      });
    }
    await app.open(book.id);
    const reopened = app.getWorkbook();
    const restored = reopened.getSheetByName('Filter limit')!;
    return {
      before,
      failures,
      states,
      reopened: {
        revision: reopened.revision,
        value: await restored.range('B1').getValues(),
        filter: reopened.metadata.sheets.find((item) => item.id === restored.id)?.filter,
        hidden: reopened.filteredRows(restored.id),
      },
    };
  });
  expect(result.failures).toEqual(['LIMIT_EXCEEDED', 'LIMIT_EXCEEDED']);
  for (const state of result.states)
    expect(state).toEqual({ ...result.before, value: [[1]], filter: null, hidden: [] });
  expect(result.reopened).toEqual({
    revision: result.before.revision,
    value: [[1]],
    filter: null,
    hidden: [],
  });
});
