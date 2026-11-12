import { expect, test } from '@playwright/test';

for (const [action, button] of [
  ['paste-special', 'Apply paste'],
  ['fill-series', 'Fill series'],
  ['annotation', 'Save note and link'],
  ['annotation', 'Remove note and link'],
  ['sort-custom', 'Apply sort'],
  ['filter-custom', 'Apply filter'],
  ['filter-custom', 'Clear all filters'],
  ['replace', 'Replace all'],
] as const) {
  for (const mode of ['reattach', 'reopen'] as const)
    test(`${action} ${button} cannot submit after ${mode}`, async ({ page }) => {
      await page.goto('/#budget');
      await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
      const before = await page.evaluate(async (clear) => {
        const app = window.opensheet;
        const book = app.getWorkbook();
        const sheet = book.getSheets()[0];
        await sheet.range('A1:A3').setValues([['replace me'], [3], [1]]);
        await sheet.range('A1').setNote('original note');
        if (clear) await sheet.setFilter({ column: 0, query: '1' });
        app.selection.set({
          sheetId: sheet.id,
          startRow: 0,
          endRow: 3,
          startColumn: 0,
          endColumn: 1,
        });
        const execute = book.execute.bind(book);
        (window as any).staleSubmissions = 0;
        book.execute = async (...args: Parameters<typeof book.execute>) => {
          (window as any).staleSubmissions++;
          return execute(...args);
        };
        return {
          revision: book.revision,
          values: await sheet.range('A1:A3').getValues(),
          notes: await sheet.range('A1').getNotes(),
          filter: book.sheetData(sheet.id).filter,
        };
      }, button === 'Clear all filters');
      await page.getByLabel('Sheet actions').selectOption(action);
      if (action === 'paste-special') {
        await page.getByLabel('Paste clipboard here').evaluate((element) => {
          const clipboard = new DataTransfer();
          clipboard.setData('text/plain', 'pasted');
          const event = new Event('paste', { bubbles: true, cancelable: true });
          Object.defineProperty(event, 'clipboardData', { value: clipboard });
          element.dispatchEvent(event);
        });
      } else if (action === 'fill-series') {
        await expect(page.getByRole('button', { name: button })).toBeEnabled();
        await page.getByLabel('Start value').fill('5');
      } else if (action === 'annotation') {
        await expect(page.getByRole('button', { name: button })).toBeEnabled();
        if (button === 'Save note and link')
          await page.getByLabel('Cell note').fill('changed note');
      } else if (action === 'replace') {
        await page.getByLabel('Find text').fill('replace');
        await page.getByLabel('Replace with').fill('changed');
      }
      await page.evaluate(async (mode) => {
        const app = window.opensheet;
        const book = app.getWorkbook();
        if (mode === 'reattach') await app.attachWorkbook(book);
        else await app.open(book.id);
      }, mode);
      await page.getByRole('button', { name: button }).click();
      await expect(page.getByRole('alert')).toHaveText(
        'Workbook changed. Check the previous operation before trying again.',
      );
      const after = await page.evaluate(async () => {
        const book = window.opensheet.getWorkbook();
        const sheet = book.getSheets()[0];
        return {
          revision: book.revision,
          values: await sheet.range('A1:A3').getValues(),
          notes: await sheet.range('A1').getNotes(),
          filter: book.sheetData(sheet.id).filter,
          calls: (window as any).staleSubmissions,
        };
      });
      expect(after).toEqual({ ...before, calls: 0 });
    });
}

test('stale paste dialog rejects clipboard capture after reopening the workbook', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.getByLabel('Sheet actions').selectOption('paste-special');
  await page.evaluate(async () => {
    const app = window.opensheet;
    await app.open(app.getWorkbook().id);
  });
  await page.getByLabel('Paste clipboard here').evaluate((element) => {
    const clipboard = new DataTransfer();
    clipboard.setData('text/plain', 'new clipboard');
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboard });
    element.dispatchEvent(event);
  });
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Check the previous operation before trying again.',
  );
  await expect(page.getByRole('button', { name: 'Apply paste' })).toBeDisabled();
  expect(errors).toEqual([]);
});
