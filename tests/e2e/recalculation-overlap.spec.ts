import { expect, test } from '@playwright/test';

for (const cacheBytes of [0, 4 * 1024 * 1024]) {
  test(`overlapping formulas preserve results and history with cache ${cacheBytes}`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(
      async ({ moduleUrl, cacheBytes }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ moduleUrl);
        let book = await createWorkbook(
          {
            sheets: [
              { name: 'Source', rows: 512, columns: 2 },
              { name: 'Output', rows: 512, columns: 2 },
            ],
          },
          { cacheBytes },
        );
        try {
          const [source, output] = book.getSheets();
          await source.range('A1:A32').setValues(Array.from({ length: 32 }, () => [1]));
          await output
            .range('A1:A512')
            .setFormulas(Array.from({ length: 512 }, (_, i) => [`SUM(Source!$A$1:$A$32)+${i}`]));
          const verify = async (expectedSource: number) => {
            const values = await book.getSheets()[1].range('A1:A512').getValues();
            for (let i = 0; i < values.length; i++)
              if (values[i][0] !== expectedSource + 31 + i)
                throw new Error(`Wrong result at ${i}: ${JSON.stringify(values[i])}`);
            return values.length;
          };
          const counts = [await verify(1)];
          await source.range('A1').setValues([[2]]);
          counts.push(await verify(2));
          const undo = await book.undo();
          counts.push(await verify(1));
          const redo = await book.redo();
          counts.push(await verify(2));
          const id = book.id;
          await book.close();
          book = await openWorkbook(id, { cacheBytes });
          counts.push(await verify(2));
          return { counts, undo, redo };
        } finally {
          await book.deleteStorage();
          book.dispose();
        }
      },
      { moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`, cacheBytes },
    );
    expect(result).toEqual({ counts: [512, 512, 512, 512, 512], undo: true, redo: true });
  });
}
