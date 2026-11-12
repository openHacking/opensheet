import { expect, test } from '@playwright/test';

test('range reads retain the row identity of their completed revision', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 2 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setFormulas([['1+2']]);
    await sheet.range('A1').setNote('old row');
    await sheet.range('A1').setLink({ target: 'https://example.com' });
    const original = book.readRange.bind(book);
    let count = 0;
    let started!: () => void;
    const waiting = new Promise<void>((resolve) => {
      started = resolve;
    });
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    book.readRange = async (...args) => {
      const data = await original(...args);
      if (++count === 6) started();
      await gate;
      return data;
    };
    const range = sheet.range('A1');
    const reads = Promise.all([
      range.getValues(),
      range.getDisplayValues(),
      range.getFormulas(),
      range.getInputs(),
      range.getNotes(),
      range.getLinks(),
    ]);
    await waiting;
    await sheet.insertRows(0);
    range.bounds.startRow = 1;
    range.bounds.endRow = 2;
    release();
    const old = await reads;
    book.readRange = original;
    const current = [await sheet.range('A1').getValues(), await sheet.range('A2').getValues()];
    return { old, current };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.old).toEqual([
    [[3]],
    [['3']],
    [['1+2']],
    [[{ type: 'formula', expression: '1+2' }]],
    [['old row']],
    [[{ target: 'https://example.com' }]],
  ]);
  expect(result.current).toEqual([[[null]], [[3]]]);
});

test('range read rejects a response from a newer axis revision', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 1 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[42]]);
    const original = book.readRange.bind(book);
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    book.readRange = async (...args) => {
      await gate;
      return original(...args);
    };
    const pending = sheet
      .range('A1')
      .getValues()
      .then(
        () => 'success',
        (error) => error.code,
      );
    await sheet.insertRows(0);
    release();
    const code = await pending;
    book.readRange = original;
    return { code, values: await sheet.range('A1:A2').getValues() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ code: 'REVISION_CONFLICT', values: [[null], [42]] });
});
