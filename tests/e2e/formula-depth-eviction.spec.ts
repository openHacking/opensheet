import { expect, test } from '@playwright/test';

test('keeps formula depth results stable before and after Worker cache eviction', async ({
  page,
}) => {
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
  expect(result.warm).toEqual({ error: '#LIMIT!' });
  expect(result.evicted).toEqual({ error: '#LIMIT!' });
  expect(result.reopened).toEqual({ error: '#LIMIT!' });
  expect(result.afterReads).toBe(result.revision);
});

for (const cacheBytes of [0, 1024 * 1024]) {
  test(`invalidates cached depth after branch and chain edits with cache ${cacheBytes}`, async ({
    page,
  }) => {
    test.setTimeout(60_000);
    await page.goto('/');
    const steps = await page.evaluate(
      async ({ url, cacheBytes }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
        let book = await createWorkbook(
          {
            sheets: [
              { name: 'Chain', rows: 130, columns: 1 },
              { name: 'Branch', rows: 2, columns: 2 },
            ],
          },
          { cacheBytes },
        );
        try {
          let [chain, branch] = book.getSheets();
          await chain.range('A1').setValues([[1]]);
          await chain
            .range('A2:A130')
            .setFormulas(Array.from({ length: 129 }, (_, row) => [`A${row + 1}+1`]));
          await branch.range('B1').setValues([[false]]);
          await branch.range('A1:A2').setFormulas([['IF(B1,Chain!A129,7)'], ['Chain!A128']]);
          const results: unknown[] = [];
          const sample = async () => {
            // Repeated reads also exercise cache hits and depth metadata.
            const first = await branch.range('A1:A2').getValues();
            const second = await branch.range('A1:A2').getValues();
            if (JSON.stringify(first) !== JSON.stringify(second)) throw new Error('Unstable read');
            results.push(first);
          };
          const reopen = async () => {
            const id = book.id;
            await book.close();
            book = await openWorkbook(id, { cacheBytes });
            [chain, branch] = book.getSheets();
          };
          await sample();
          await branch.range('B1').setValues([[true]]);
          await sample();
          await chain.range('A100').setValues([[5]]);
          await sample();
          if (!(await book.undo())) throw new Error('Undo chain edit failed');
          await sample();
          if (!(await book.undo())) throw new Error('Undo branch edit failed');
          await sample();
          if (!(await book.redo())) throw new Error('Redo branch edit failed');
          await sample();
          if (!(await book.redo())) throw new Error('Redo chain edit failed');
          await sample();
          await reopen();
          await sample();
          await chain.range('A100').setFormulas([['A99+1']]);
          await sample();
          if (!(await book.undo())) throw new Error('Undo restored formula failed');
          await sample();
          await reopen();
          await sample();
          return results;
        } finally {
          await book.deleteStorage();
        }
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, cacheBytes },
    );
    const limit = { error: '#LIMIT!' };
    expect(steps).toEqual([
      [[7], [128]],
      [[limit], [128]],
      [[34], [33]],
      [[limit], [128]],
      [[7], [128]],
      [[limit], [128]],
      [[34], [33]],
      [[34], [33]],
      [[limit], [128]],
      [[34], [33]],
      [[34], [33]],
    ]);
  });
}
