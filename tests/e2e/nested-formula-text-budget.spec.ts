import { expect, test } from '@playwright/test';
import { resolve } from 'node:path';

test('nested conditional and lookup work shares a per-expression budget', async ({ page }) => {
  await page.goto('/');
  const results = await page.evaluate(
    async (url) => {
      const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
      const output = [];
      for (const cacheBytes of [0, 1024 * 1024]) {
        let book = await createWorkbook(
          { sheets: [{ name: 'Data', rows: 25, columns: 3 }] },
          { cacheBytes },
        );
        try {
          let sheet = book.getSheets()[0];
          const text = 'x'.repeat(32767);
          await sheet.range('A1:A20').setValues(Array.from({ length: 20 }, () => [text]));
          const count = 'COUNTIF(A1:A20,"needle")';
          const lookup = 'XLOOKUP("needle",A1:A20,A1:A20,0,0)';
          await sheet
            .range('B1:B6')
            .setFormulas([
              [count],
              [`SUM(${count},${count})`],
              [`IFERROR(SUM(${count},${count}),77)`],
              [`IF(FALSE(),SUM(${count},${count}),7)`],
              ['SUM(COUNTIF(A1:A15,"needle"),COUNTIF(A1:A15,"needle"))'],
              [`SUM(${count},${lookup})`],
            ]);
          const initial = await sheet.range('B1:B6').getValues();
          await sheet.range('A15:A20').setValues(Array.from({ length: 6 }, () => ['x']));
          const recovered = await sheet.range('B1:B6').getValues();
          await book.undo();
          const undone = await sheet.range('B1:B6').getValues();
          await book.redo();
          const id = book.id;
          await book.close();
          book = await openWorkbook(id, { cacheBytes });
          sheet = book.getSheets()[0];
          const reopened = await sheet.range('B1:B6').getValues();
          output.push({
            cacheBytes,
            initial,
            recovered,
            undone,
            reopened,
            originalLength: (await sheet.range('A1').getValues())[0][0].length,
          });
        } finally {
          await book.close();
        }
      }
      return output;
    },
    `/@fs/${resolve('packages/core/src/index.ts')}`,
  );
  const limited = [
    [0],
    [{ error: '#LIMIT!' }],
    [{ error: '#LIMIT!' }],
    [7],
    [0],
    [{ error: '#LIMIT!' }],
  ];
  const recovered = [[0], [0], [0], [7], [0], [0]];
  expect(results).toEqual(
    [0, 1024 * 1024].map((cacheBytes) => ({
      cacheBytes,
      initial: limited,
      recovered,
      undone: limited,
      reopened: recovered,
      originalLength: 32767,
    })),
  );
});
