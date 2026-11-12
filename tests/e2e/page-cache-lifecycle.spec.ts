import { expect, test } from '@playwright/test';

test('a page-cache transition retains the editor and an unfinished draft', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    window.opensheet.selection.set({
      sheetId: sheet.id,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
    return { id: book.id, revision: book.revision };
  });
  await page.getByRole('grid', { name: 'Spreadsheet' }).focus();
  await page.keyboard.type('retained draft');
  const editor = page.getByRole('textbox', { name: 'Edit cell' });
  await expect(editor).toHaveValue('retained draft');
  const result = await page.evaluate(() => {
    const before = new Event('beforeunload', { cancelable: true });
    window.dispatchEvent(before);
    window.dispatchEvent(new PageTransitionEvent('pagehide', { persisted: true }));
    window.dispatchEvent(new PageTransitionEvent('pageshow', { persisted: true }));
    return {
      warned: before.defaultPrevented,
      draft: window.opensheet.hasDraft,
      id: window.opensheet.getWorkbook().id,
    };
  });
  expect(result).toEqual({ warned: true, draft: true, id: initial.id });
  await expect(editor).toHaveValue('retained draft');
  await editor.press('Enter');
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(initial.revision + 1);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([['retained draft']]);
  await page.evaluate(() =>
    window.dispatchEvent(new PageTransitionEvent('pagehide', { persisted: false })),
  );
  expect(
    await page.evaluate(() => {
      try {
        window.opensheet.getWorkbook();
        return false;
      } catch {
        return true;
      }
    }),
  ).toBe(true);
});

test('browser back navigation restores a usable workbook', async ({ page }, info) => {
  await page.addInitScript(() => {
    (window as any).cacheShows = [];
    window.addEventListener('pageshow', (event) =>
      (window as any).cacheShows.push(event.persisted),
    );
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const initial = await page.evaluate(() => {
    (window as any).cacheIdentity = crypto.randomUUID();
    return { id: window.opensheet.getWorkbook().id, identity: (window as any).cacheIdentity };
  });
  await page.goto('about:blank');
  await page.goBack();
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['after return']]);
    return {
      id: book.id,
      identity: (window as any).cacheIdentity,
      shows: (window as any).cacheShows,
      value: await book.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result.id).toBe(initial.id);
  expect(result.value).toEqual([['after return']]);
  const restoredFromCache = result.shows.at(-1) === true;
  expect(result.identity === initial.identity).toBe(restoredFromCache);
  info.annotations.push({
    type: 'page cache',
    description: restoredFromCache ? 'restored' : 'not used by browser',
  });
});
