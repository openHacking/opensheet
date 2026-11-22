import { expect, test } from '@playwright/test';

test('copies wide ranges in bounded batches and rejects a mid-copy edit', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ coreUrl, clipboardUrl, structuredUrl }) => {
      const { createWorkbook } = await import(/* @vite-ignore */ coreUrl);
      const { copyText, cutText } = await import(/* @vite-ignore */ clipboardUrl);
      const { copyClipboard } = await import(/* @vite-ignore */ structuredUrl);
      const book = await createWorkbook({ sheets: [{ name: 'Wide', rows: 3, columns: 300 }] });
      try {
        const sheet = book.getSheets()[0];
        const rows = Array.from({ length: 3 }, (_, row) =>
          Array.from({ length: 300 }, (_, column) => `${row}:${column}`),
        );
        await sheet.range('A1:KN3').setValues(rows);
        const source = sheet.range('A1:KN3');
        const counts: number[] = [];
        const readRange = book.readRange.bind(book);
        const readCells = book.readCells.bind(book);
        book.readRange = async (
          id: string,
          bounds: { startRow: number; endRow: number; startColumn: number; endColumn: number },
          values: boolean,
        ) => {
          counts.push((bounds.endRow - bounds.startRow) * (bounds.endColumn - bounds.startColumn));
          return readRange(id, bounds, values);
        };
        book.readCells = async (
          id: string,
          bounds: { startRow: number; endRow: number; startColumn: number; endColumn: number },
          options: unknown,
        ) => {
          counts.push((bounds.endRow - bounds.startRow) * (bounds.endColumn - bounds.startColumn));
          return readCells(id, bounds, options);
        };
        const plain = await copyText(source);
        const structured = await copyClipboard(source);
        let reads = 0,
          writes = 0;
        book.readRange = async (...args: any[]) => {
          if (++reads === 2) await sheet.range('A3').setValues([['changed']]);
          return readRange(...args);
        };
        const code = await cutText(book, source, async () => {
          writes++;
        }).then(
          () => 'success',
          (error: { code: string }) => error.code,
        );
        book.readRange = readRange;
        const values = await sheet.range('A1:A3').getValues();
        const id = book.id;
        await book.close();
        const { openWorkbook } = await import(/* @vite-ignore */ coreUrl);
        const reopened = await openWorkbook(id);
        try {
          return {
            rows,
            plain,
            structured: structured.text,
            counts,
            code,
            writes,
            values,
            reopened: await reopened.getSheets()[0].range('A1:A3').getValues(),
          };
        } finally {
          await reopened.deleteStorage();
        }
      } finally {
        book.dispose();
      }
    },
    {
      coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      clipboardUrl: `/@fs${process.cwd()}/packages/renderer/src/clipboard.ts`,
      structuredUrl: `/@fs${process.cwd()}/packages/renderer/src/copy-clipboard.ts`,
    },
  );
  const expected = result.rows.map((row: string[]) => row.join('\t')).join('\n');
  expect(result.plain).toBe(expected);
  expect(result.structured).toBe(expected);
  expect(result.counts.length).toBeGreaterThan(3);
  expect(Math.max(...result.counts)).toBeLessThanOrEqual(256);
  expect(result.code).toBe('REVISION_CONFLICT');
  expect(result.writes).toBe(0);
  expect(result.values).toEqual([['0:0'], ['1:0'], ['changed']]);
  expect(result.reopened).toEqual(result.values);
});

test('rejects oversized quoted plain-text cut without writing or clearing', async ({ page }) => {
  test.setTimeout(60_000);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ coreUrl, clipboardUrl }) => {
      const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ coreUrl);
      const { cutText } = await import(/* @vite-ignore */ clipboardUrl);
      let book = await createWorkbook({ sheets: [{ name: 'Long text', rows: 520, columns: 1 }] });
      try {
        let sheet = book.getSheets()[0];
        const value = '"'.repeat(32768);
        await sheet.range('A1:A520').setValues(Array.from({ length: 520 }, () => [value]));
        const revision = book.revision;
        let writes = 0;
        const error = await cutText(book, sheet.range('A1:A520'), async () => {
          writes++;
        }).then(
          () => 'success',
          (error: { code: string }) => error.code,
        );
        const after = book.revision;
        const first = (await sheet.range('A1').getValues())[0][0];
        const last = (await sheet.range('A520').getValues())[0][0];
        const id = book.id;
        await book.close();
        book = await openWorkbook(id);
        sheet = book.getSheets()[0];
        return {
          error,
          writes,
          revision,
          after,
          reopenedRevision: book.revision,
          first: first === value,
          last: last === value,
          reopened: (await sheet.range('A520').getValues())[0][0] === value,
        };
      } finally {
        await book.deleteStorage();
      }
    },
    {
      coreUrl: `/@fs${process.cwd()}/packages/core/src/index.ts`,
      clipboardUrl: `/@fs${process.cwd()}/packages/renderer/src/clipboard.ts`,
    },
  );
  expect(result).toEqual({
    error: 'LIMIT_EXCEEDED',
    writes: 0,
    revision: result.revision,
    after: result.revision,
    reopenedRevision: result.revision,
    first: true,
    last: true,
    reopened: true,
  });
});
