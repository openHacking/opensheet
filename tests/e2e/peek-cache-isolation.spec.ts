import { expect, test } from '@playwright/test';

test('renderer cache peeks cannot mutate displayed or persisted cell state', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('A1').setNote('original');
    await sheet.range('A1').setLink({ target: 'https://example.com' });
    await sheet.range('B1').setFormulas([['1/0']]);
    await book.prefetch(sheet.id, sheet.range('A1:B1').bounds);
    const cell = book.peekCell(sheet.id, 0, 0)!;
    const error = book.peekValue(sheet.id, 0, 1) as { error: string };
    const mutations = [
      Reflect.set(cell.input, 'value', 99),
      Reflect.set(cell, 'note', 'changed'),
      Reflect.set(cell.link!, 'target', 'https://other.example'),
      Reflect.set(error, 'error', '#N/A'),
    ];
    const cached = {
      cell: book.peekCell(sheet.id, 0, 0),
      value: book.peekValue(sheet.id, 0, 0),
      error: book.peekValue(sheet.id, 0, 1),
      display: book.peekDisplay(sheet.id, 0, 0),
    };
    const live = {
      values: await sheet.range('A1:B1').getValues(),
      note: await sheet.range('A1').getNotes(),
      link: await sheet.range('A1').getLinks(),
    };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return {
      mutations,
      cached,
      live,
      reopened: {
        values: await sheet.range('A1:B1').getValues(),
        note: await sheet.range('A1').getNotes(),
        link: await sheet.range('A1').getLinks(),
      },
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.mutations).toEqual([false, false, false, false]);
  expect(result.cached.cell).toMatchObject({
    input: { type: 'number', value: 7 },
    note: 'original',
    link: { target: 'https://example.com' },
  });
  expect(result.cached).toMatchObject({ value: 7, error: { error: '#DIV/0!' }, display: '7' });
  expect(result.live).toEqual({
    values: [[7, { error: '#DIV/0!' }]],
    note: [['original']],
    link: [[{ target: 'https://example.com' }]],
  });
  expect(result.reopened).toEqual(result.live);
});
