import { expect, test } from '@playwright/test';

test('single-cell reads use identities from their completed revision', async ({ page }) => {
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
    let count = 0;
    let ready!: () => void;
    const waiting = new Promise<void>((resolve) => {
      ready = resolve;
    });
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    book.readRange = async (...args) => {
      const data = await original(...args);
      if (++count === 3) ready();
      await gate;
      return data;
    };
    const reads = Promise.all([
      book.getCell(sheet.id, 0, 0),
      book.value(sheet.id, 0, 0),
      book.display(sheet.id, 0, 0),
    ]);
    await waiting;
    await sheet.insertRows(0);
    release();
    const [cell, value, display] = await reads;
    book.readRange = original;
    return { input: cell?.input, value, display, now: await sheet.range('A1:A2').getValues() };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    input: { type: 'number', value: 42 },
    value: 42,
    display: '42',
    now: [[null], [42]],
  });
});

test('streamed reads retain their initial range and value option across batches', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 20000, columns: 1 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:A20000').fillSeries({ start: 1, step: 1 });
    const range = { startRow: 0, endRow: 20000, startColumn: 0, endColumn: 1 };
    const options = { values: true };
    const stream = book.streamRange(sheet.id, range, options);
    const first = await stream.next();
    range.endRow = 16385;
    options.values = false;
    const second = await stream.next();
    const end = await stream.next();
    return {
      first: first.value?.range,
      second: second.value?.range,
      last: second.value?.calculated['19999:0'],
      done: end.done,
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    first: { startRow: 0, endRow: 16384, startColumn: 0, endColumn: 1 },
    second: { startRow: 16384, endRow: 20000, startColumn: 0, endColumn: 1 },
    last: 20000,
    done: true,
  });
});

test('copy special retains its original source when range bounds are edited during the read', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async (url) => {
    const { createWorkbookFile } = await import(/* @vite-ignore */ url);
    const app = window.opensheet;
    await app.load(createWorkbookFile({ sheets: [{ name: 'Sheet', rows: 3, columns: 4 }] }));
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1:B1').setValues([[1, 2]]);
    const source = sheet.range('A1:B1');
    const original = book.readCells.bind(book);
    let ready!: () => void;
    const waiting = new Promise<void>((resolve) => {
      ready = resolve;
    });
    let release!: () => void;
    const gate = new Promise<void>((resolve) => {
      release = resolve;
    });
    book.readCells = async (...args) => {
      const data = await original(...args);
      ready();
      await gate;
      return data;
    };
    const pending = source.copyTo('C1:D1', { mode: 'values' });
    await waiting;
    source.bounds.startColumn = 1;
    source.bounds.endColumn = 3;
    release();
    await pending;
    book.readCells = original;
    return sheet.range('A1:D1').getValues();
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual([[1, 2, 1, 2]]);
});
