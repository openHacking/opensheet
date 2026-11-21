import { readFileSync } from 'node:fs';
import { expect, test } from '@playwright/test';

const corpus = JSON.parse(
  readFileSync(new URL('../fixtures/lookup-error-corpus.json', import.meta.url), 'utf8'),
);

for (const cacheBytes of [0, 65536]) {
  test(`lookup error and blank corpus survives edit, history and reopen with cache ${cacheBytes}`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(
      async ({ url, corpus, cacheBytes }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
        let book = await createWorkbook(
          { sheets: [{ name: 'Lookup', rows: 20, columns: 5 }] },
          { cacheBytes },
        );
        try {
          let sheet = book.getSheets()[0];
          await sheet.range('A1:A6').setValues([[1], [null], [3], [null], [null], [null]]);
          await book.execute({
            type: 'core.cells.set',
            payload: {
              sheetId: sheet.id,
              cells: [{ row: 1, column: 0, input: { type: 'error', code: '#DIV/0!' } }],
            },
          });
          await sheet.range('A6').setFormulas([['""']]);
          await sheet.range('B1:B6').setValues(corpus.results.map((value: number) => [value]));
          await sheet
            .range(`D1:D${corpus.cases.length}`)
            .setFormulas(corpus.cases.map((item: { formula: string }) => [item.formula]));
          const read = () => sheet.range(`D1:D${corpus.cases.length}`).getValues();
          const initial = await read();
          await sheet.range('A2').setValues([[2]]);
          const edited = await read();
          await book.undo();
          const undone = await read();
          await book.redo();
          const redone = await read();
          await book.undo();
          const revision = book.revision;
          const id = book.id;
          await book.close();
          book = await openWorkbook(id, { cacheBytes });
          sheet = book.getSheets()[0];
          return {
            initial,
            edited,
            undone,
            redone,
            reopened: await read(),
            revision,
            reopenedRevision: book.revision,
          };
        } finally {
          await book.deleteStorage();
        }
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, corpus, cacheBytes },
    );
    const expected = corpus.cases.map((item: { expected: unknown }) => [item.expected]);
    expect(result.initial).toEqual(expected);
    expect(result.undone).toEqual(expected);
    expect(result.reopened).toEqual(expected);
    expect(result.reopenedRevision).toBe(result.revision);
    expect(result.edited.slice(0, 5)).toEqual([[3], [1], [2], [30], [10]]);
    expect(result.redone).toEqual(result.edited);
  });
}
