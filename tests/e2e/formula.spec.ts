import { test, expect } from '@playwright/test';
import {
  formulaCorpus,
  date1904Corpus,
  formulaLookupInputs,
  orderedLookupInputs,
} from '../fixtures/formula-corpus.js';
import { createWorkbookFile } from '@opensheetjs/core';

test('decimal rounding survives real Worker calculation, edit and reopen', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[1.005]]);
    await sheet.range('B1:C1').setFormulas([['ROUND(A1,2)', 'ROUND(1,101)']]);
    const first = await sheet.range('B1:C1').getValues();
    await sheet.range('A1').setValues([[-1.475]]);
    const changed = await sheet.range('B1:C1').getValues();
    const id = book.id;
    await app.open(id);
    book = app.getWorkbook();
    return { first, changed, reopened: await book.getSheets()[0].range('B1:C1').getValues() };
  });
  expect(result).toEqual({ first: [[1.01, 1]], changed: [[-1.48, 1]], reopened: [[-1.48, 1]] });
});

test('formula-key external sorting preserves stable rows and persistent history', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async (fileJSON) => {
      const app = window.opensheet;
      await app.load(JSON.parse(fileJSON));
      const book = app.getWorkbook(),
        sheet = book.getSheets()[0];
      await sheet.range('A1:A600').setValues(Array.from({ length: 600 }, (_, i) => [i]));
      await sheet
        .range('B1:B600')
        .setFormulas(Array.from({ length: 600 }, (_, i) => [`MOD(A${i + 1}+7,20)`]));
      const start = performance.now();
      await sheet.sort('A1:B600', 1);
      const sortMs = performance.now() - start;
      const values = await sheet.range('A1:B600').getValues();
      await book.undo();
      const undone = await sheet.range('A1:A3').getValues();
      await book.redo();
      await app.open(book.id);
      const restored = await app.getWorkbook().getSheets()[0].range('A1:B600').getValues();
      const reopenedSheet = app.getWorkbook().getSheets()[0];
      const multiStart = performance.now();
      await reopenedSheet.sort('A1:B600', [
        { column: 1, direction: 'asc' },
        { column: 0, direction: 'desc' },
      ]);
      const multiSortMs = performance.now() - multiStart;
      const multi = await reopenedSheet.range('A1:B600').getValues();
      await app.getWorkbook().undo();
      const multiUndone = await reopenedSheet.range('A1:B600').getValues();
      return { values, undone, restored, sortMs, multi, multiSortMs, multiUndone };
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Sort', rows: 600 }] })),
  );
  const expected = Array.from({ length: 600 }, (_, i) => [i, (i + 7) % 20]).sort(
    (a, b) => a[1] - b[1] || a[0] - b[0],
  );
  expect(result.values).toEqual(expected);
  expect(result.restored).toEqual(expected);
  expect(result.undone).toEqual([[0], [1], [2]]);
  expect(result.multi).toEqual([...expected].sort((a, b) => a[1] - b[1] || b[0] - a[0]));
  expect(result.multiUndone).toEqual(expected);
  console.log(`600-row multi-key formula sort: ${result.multiSortMs.toFixed(1)} ms`);
  console.log(`600-row formula sort: ${result.sortMs.toFixed(1)} ms`);
});

for (const dateSystem of ['1900', '1904'] as const) {
  const corpus = dateSystem === '1900' ? formulaCorpus : date1904Corpus;
  test(`formula semantics survive real Worker storage and reopen (${dateSystem})`, async ({
    page,
  }) => {
    await page.goto('/');
    await page.waitForFunction(() => !!window.opensheet?.selection.get());
    await expect(page.locator('#scene-select')).toBeEnabled();
    const expressions = corpus.map(([expression]) => expression);
    const values = await page.evaluate(
      async ({ expressions, fileJSON, lookupInputs, orderedInputs }) => {
        const app = window.opensheet;
        await app.ready();
        await app.load(JSON.parse(fileJSON));
        const book = app.getWorkbook();
        const sheet = book.getSheets()[0];
        await sheet.range('A1:A7').setValues([[true], ['7'], [2], [null], [''], [null], [false]]);
        await sheet.range('D1:E5').setValues(lookupInputs);
        await sheet.range('G1:I5').setValues(orderedInputs);
        await sheet.range('A6').setFormulas([['1/0']]);
        await sheet
          .range(`B1:B${expressions.length}`)
          .setFormulas(expressions.map((expression) => [expression]));
        const initial = await sheet.range(`B1:B${expressions.length}`).getValues();
        await app.open(book.id);
        const restored = await app
          .getWorkbook()
          .getSheets()[0]
          .range(`B1:B${expressions.length}`)
          .getValues();
        return { initial, restored };
      },
      {
        expressions,
        lookupInputs: formulaLookupInputs,
        orderedInputs: orderedLookupInputs,
        fileJSON: JSON.stringify({
          ...createWorkbookFile({
            sheets: [{ name: 'Semantics', rows: Math.max(100, corpus.length) }],
          }),
          dateSystem,
        }),
      },
    );
    const expected = corpus.map(([, value]) => [value]);
    expect(values.initial).toEqual(expected);
    expect(values.restored).toEqual(expected);
  });
}
