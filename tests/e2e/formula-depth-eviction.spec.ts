import { expect, test } from '@playwright/test';

test('documents formula depth results before and after Worker cache eviction', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
    let book = await createWorkbook(
      {
        sheets: [
          { name: 'Chain', rows: 130, columns: 1 },
          { name: 'Filler', rows: 2000, columns: 1 },
        ],
      },
      { cacheBytes: 1024 * 1024 },
    );
    try {
      let [chain, filler] = book.getSheets();
      await chain.range('A1').setValues([[1]]);
      await chain
        .range('A2:A130')
        .setFormulas(Array.from({ length: 129 }, (_, row) => [`A${row + 1}+1`]));
      await filler
        .range('A1:A2000')
        .setFormulas(Array.from({ length: 2000 }, (_, row) => [`${row}+1`]));
      const revision = book.revision;
      const read = async (address: string) => (await chain.range(address).getValues())[0][0];
      const cold = await read('A130');
      const intermediate = await read('A129');
      const warm = await read('A130');
      const values = await filler.range('A1:A2000').getValues();
      if (values.some((row: unknown[], index: number) => row[0] !== index + 1))
        throw new Error('Filler calculation failed');
      const evicted = await read('A130');
      const afterReads = book.revision;
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 1024 * 1024 });
      [chain, filler] = book.getSheets();
      const reopened = await read('A130');
      return { cold, intermediate, warm, evicted, reopened, revision, afterReads };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.cold).toEqual({ error: '#LIMIT!' });
  expect(result.intermediate).toBe(129);
  expect(result.warm).toBe(130);
  expect(result.evicted).toEqual({ error: '#LIMIT!' });
  expect(result.reopened).toEqual({ error: '#LIMIT!' });
  expect(result.afterReads).toBe(result.revision);
});
