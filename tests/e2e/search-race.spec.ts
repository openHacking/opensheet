import { workbookTool } from './workbook-tools.js';
import { expect, test } from '@playwright/test';

test('search uses current row identities after insertion, undo and reopen', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([['needle'], ['other'], ['needle']]);
    const before = await book.search(sheet.id, 'needle');
    await sheet.insertRows(0, 2);
    const inserted = await book.search(sheet.id, 'needle');
    await book.undo();
    const undone = await book.search(sheet.id, 'needle');
    await book.redo();
    const id = book.id;
    await app.open(id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await book.search(sheet.id, 'needle');
    return { before, inserted, undone, reopened };
  });
  expect(result).toEqual({
    before: { count: 2, match: { row: 0, column: 0 } },
    inserted: { count: 2, match: { row: 2, column: 0 } },
    undone: { count: 2, match: { row: 0, column: 0 } },
    reopened: { count: 2, match: { row: 2, column: 0 } },
  });
});

test('finds formatted values and clearly replaces only case-sensitive literal text', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    await sheet.range('A1:A2').setValues([['Alpha alpha'], [42]]);
    await sheet.range('A2').setStyle({ numberFormat: '0.00' });
    await sheet.range('B1').setFormulas([['"Alpha"']]);
  });
  await page.locator('#find').fill('42.00');
  await page.locator('#find-next').click();
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(1);
  await page.locator('#find').fill('Alpha');
  await workbookTool(page, 'replace');
  await expect(page.getByRole('dialog')).toContainText('case-sensitive literal text');
  await page.getByLabel('Replace with').fill('Omega');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  await expect
    .poll(() =>
      page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues()),
    )
    .toEqual([['Omega alpha', 'Alpha']]);
  await page.evaluate(() => window.opensheet.getWorkbook().undo());
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:B1').getValues(),
    ),
  ).toEqual([['Alpha alpha', 'Alpha']]);
});

test('failed oversized replacement preserves both pages and undo history in the real Worker', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const longText = 'a'.repeat(32768);
    await sheet.range('A1').setValues([['a']]);
    await sheet.range('A100').setValues([[longText]]);
    const revision = book.revision;
    let code = '';
    try {
      await book.replaceText(sheet.id, 'a', 'x'.repeat(32768));
    } catch (error) {
      code = (error as { code?: string }).code ?? '';
    }
    const after = {
      revision: book.revision,
      first: await sheet.range('A1').getValues(),
      last: await sheet.range('A100').getValues(),
    };
    const undo = await book.undo();
    const undone = await sheet.range('A100').getValues();
    return { code, revision, after, undo, undone, longText };
  });
  expect(result).toMatchObject({
    code: 'LIMIT_EXCEEDED',
    after: { revision: result.revision, first: [['a']], last: [[result.longText]] },
    undo: true,
    undone: [[null]],
  });
});

test('find advances and wraps through real Worker results', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1:A3')
      .setValues([['needle'], ['other'], ['needle']]);
  });
  await page.locator('#find').fill('needle');
  for (const row of [0, 2, 0]) {
    await page.locator('#find-next').click();
    await expect
      .poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow))
      .toBe(row);
    await expect(page.locator('#notice')).toContainText('2 matching cells');
  }
  await page.getByRole('button', { name: 'Find previous' }).click();
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(2);
  await page.locator('#find').press('Shift+Enter');
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(0);
  expect(
    await page.evaluate(() => {
      const book = window.opensheet.getWorkbook();
      const original = book.search.bind(book);
      let calls = 0;
      book.search = (...args) => {
        calls++;
        return original(...args);
      };
      document.querySelector('#find')!.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'Enter',
          shiftKey: true,
          isComposing: true,
          bubbles: true,
        }),
      );
      return calls;
    }),
  ).toBe(0);
});

test('find ignores stale results and resets its cursor on sheet or selection changes', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const second = await book.addSheet('Search second');
    const pending: any[] = [];
    (window as any).searchPending = pending;
    (window as any).secondSheet = second.id;
    book.search = (sheet, query, after, options) =>
      new Promise((resolve) =>
        pending.push({ sheet, query, after, signal: options?.signal, resolve }),
      );
  });
  const initial = await page.evaluate(() => window.opensheet.selection.get());
  await page.locator('#find').fill('old');
  await page.locator('#find-next').click();
  await expect.poll(() => page.evaluate(() => (window as any).searchPending.length)).toBe(1);
  await page.locator('#find').fill('new');
  expect(await page.evaluate(() => (window as any).searchPending[0].signal.aborted)).toBe(true);
  await page.locator('#find-next').click();
  await expect.poll(() => page.evaluate(() => (window as any).searchPending.length)).toBe(2);
  await page.evaluate(() =>
    (window as any).searchPending[1].resolve({ count: 1, match: { row: 4, column: 2 } }),
  );
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(4);
  await page.evaluate(() =>
    (window as any).searchPending[0].resolve({ count: 1, match: { row: 7, column: 3 } }),
  );
  await expect.poll(() => page.evaluate(() => window.opensheet.selection.get()?.startRow)).toBe(4);
  await page.locator('#find-next').click();
  await expect.poll(() => page.evaluate(() => (window as any).searchPending.length)).toBe(3);
  expect(await page.evaluate(() => (window as any).searchPending[2].after)).toBeGreaterThanOrEqual(
    0,
  );
  await page.evaluate(() => {
    const selection = window.opensheet.selection.get()!;
    window.opensheet.selection.set({
      ...selection,
      startRow: 1,
      endRow: 2,
      startColumn: 0,
      endColumn: 1,
    });
    (window as any).searchPending[2].resolve({ count: 1, match: { row: 8, column: 2 } });
  });
  expect(await page.evaluate(() => (window as any).searchPending[2].signal.aborted)).toBe(true);
  await page.locator('#find-next').click();
  await expect.poll(() => page.evaluate(() => (window as any).searchPending.length)).toBe(4);
  expect(await page.evaluate(() => (window as any).searchPending[3].after)).toBe(-1);
  await page.evaluate(() => {
    const sheetId = (window as any).secondSheet;
    window.opensheet.selection.set({
      sheetId,
      startRow: 0,
      endRow: 1,
      startColumn: 0,
      endColumn: 1,
    });
    (window as any).searchPending[3].resolve({ count: 1, match: { row: 9, column: 2 } });
  });
  expect(await page.evaluate(() => (window as any).searchPending[3].signal.aborted)).toBe(true);
  await page.locator('#find-next').click();
  await expect.poll(() => page.evaluate(() => (window as any).searchPending.length)).toBe(5);
  expect(await page.evaluate(() => (window as any).searchPending[4].after)).toBe(-1);
  expect(await page.evaluate(() => window.opensheet.selection.get()?.sheetId)).not.toBe(
    initial?.sheetId,
  );
  await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheetById((window as any).secondSheet)!
      .range('B1')
      .setValues([['edit']]);
  });
  expect(await page.evaluate(() => (window as any).searchPending[4].signal.aborted)).toBe(true);
  const selection = await page.evaluate(() => window.opensheet.selection.get());
  await page.locator('#find-next').click();
  await expect(page.locator('.search-wrap')).toHaveAttribute('aria-busy', 'true');
  await expect(page.locator('#notice')).toContainText('Searching');
  await expect(page.locator('#find')).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.locator('.search-wrap')).not.toHaveAttribute('aria-busy', 'true');
  await expect(page.locator('#notice')).toContainText('Search cancelled');
  expect(await page.evaluate(() => (window as any).searchPending[5].signal.aborted)).toBe(true);
  await page.evaluate(() =>
    (window as any).searchPending[5].resolve({ count: 1, match: { row: 3, column: 2 } }),
  );
  expect(await page.evaluate(() => window.opensheet.selection.get())).toEqual(selection);
});
