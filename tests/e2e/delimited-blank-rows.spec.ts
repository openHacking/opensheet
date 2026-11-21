import { expect, test } from '@playwright/test';

test('CSV and TSV retain BOM text and blank rows across export batches and reopen', async ({
  page,
}) => {
  test.setTimeout(60_000);
  await page.goto('/');
  const result = await page.evaluate(
    async ({ coreUrl, formatUrl }) => {
      const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ coreUrl);
      const { exportRange, parseDelimited, streamExport } = await import(
        /* @vite-ignore */ formatUrl
      );
      let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 16400, columns: 1 }] });
      try {
        let sheet = book.getSheets()[0];
        await sheet.range('A1').setValues([['\uFEFFfirst']]);
        await sheet.range('A16385').setValues([['\uFEFFmiddle']]);
        const verify = async () => {
          const file = await book.toJSON();
          const results = [];
          for (const format of ['csv', 'tsv'] as const)
            for (const bom of [false, true]) {
              const config = { sheetId: sheet.id, format, range: 'A1:A16400', options: { bom } };
              const snapshot = exportRange(file, config).text;
              const chunks: string[] = [];
              for await (const chunk of streamExport(book, config)) chunks.push(chunk);
              const text = chunks.join('');
              const parsed = parseDelimited(text, format === 'csv' ? ',' : '\t');
              results.push({
                same: text === snapshot,
                chunks: chunks.length,
                count: parsed.length,
                first: parsed[0],
                remainingCorrect: parsed
                  .slice(1)
                  .every(
                    (row: string[], index: number) =>
                      row.length === 1 && row[0] === (index === 16383 ? '\uFEFFmiddle' : ''),
                  ),
                quotedLast: text.endsWith('""'),
              });
            }
          return results;
        };
        const before = await verify();
        const id = book.id;
        await book.close();
        book = await openWorkbook(id);
        sheet = book.getSheets()[0];
        const after = await verify();
        return { before, after };
      } finally {
        await book.deleteStorage();
      }
    },
    {
      coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      formatUrl: `/@fs${process.cwd()}/packages/formats/src/index.ts`,
    },
  );
  for (const results of [result.before, result.after]) {
    for (const item of results) {
      expect(item).toMatchObject({
        same: true,
        count: 16400,
        first: ['\uFEFFfirst'],
        remainingCorrect: true,
        quotedLast: true,
      });
      expect(item.chunks).toBeGreaterThan(1);
    }
  }
});
