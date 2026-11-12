import { expect, test } from '@playwright/test';

for (const cacheBytes of [0, 4 * 1024 * 1024]) {
  test(`generated formulas recalculate through history and reopen with cache ${cacheBytes}`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(
      async ({ moduleUrl, cacheBytes }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ moduleUrl);
        let book = await createWorkbook(
          { sheets: [{ name: 'Data', rows: 128, columns: 6 }] },
          { cacheBytes },
        );
        const random = (() => {
          let seed = 20261112;
          return () => {
            seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
            return seed;
          };
        })();
        const initial = Array.from({ length: 128 }, () => [
          (random() % 101) - 50,
          (random() % 101) - 50,
          (random() % 101) - 50,
        ]);
        const changed = initial.map(([a, b, c]) => [-a + 1, b, c]);
        const formulas = initial.map((_, i) => {
          const n = i + 1;
          return [`A${n}+B${n}*C${n}`, `IF(A${n}>0,B${n},C${n})`, `IFERROR(A${n}/0,B${n})`];
        });
        try {
          let sheet = book.getSheets()[0];
          await sheet.range('A1:C128').setValues(initial);
          await sheet.range('D1:F128').setFormulas(formulas);
          const verify = async (inputs: number[][]) => {
            const actual = await book.getSheets()[0].range('D1:F128').getValues();
            for (let i = 0; i < inputs.length; i++) {
              const [a, b, c] = inputs[i],
                expected = [a + b * c, a > 0 ? b : c, b];
              if (actual[i].some((value: unknown, column: number) => value !== expected[column]))
                throw new Error(`Formula mismatch at row ${i + 1}: ${JSON.stringify(actual[i])}`);
            }
            return actual.length * 3;
          };
          const counts = [await verify(initial)];
          await sheet.range('A1:A128').setValues(changed.map(([a]) => [a]));
          counts.push(await verify(changed));
          const undo = await book.undo();
          counts.push(await verify(initial));
          const redo = await book.redo();
          counts.push(await verify(changed));
          const id = book.id;
          await book.close();
          book = await openWorkbook(id, { cacheBytes });
          counts.push(await verify(changed));
          return { counts, undo, redo };
        } finally {
          await book.deleteStorage();
          book.dispose();
        }
      },
      { moduleUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`, cacheBytes },
    );
    expect(result).toEqual({ counts: [384, 384, 384, 384, 384], undo: true, redo: true });
  });
}
