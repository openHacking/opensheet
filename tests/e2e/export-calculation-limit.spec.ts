import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('value exports reject cold calculation limits while formula-preserving CSV succeeds', async ({
  page,
}) => {
  const formatsModule = `/@fs${fileURLToPath(new URL('../../packages/formats/src/index.ts', import.meta.url))}`;
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (formatsModule) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const sheet = await book.addSheet('Export limit', { rows: 130, columns: 1 });
    await sheet.range('A1').setValues([[1]]);
    await sheet
      .range('A2:A130')
      .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 1}+1`]));
    const id = book.id;
    const { streamExport } = await import(/* @vite-ignore */ formatsModule);
    const failures = [];
    for (const format of ['csv', 'tsv', 'html', 'markdown', 'latex']) {
      await app.open(id);
      book = app.getWorkbook();
      const target = book.getSheetByName('Export limit')!;
      const chunks: string[] = [];
      let code = '',
        reports = 0;
      try {
        for await (const chunk of streamExport(
          book,
          { sheetId: target.id, range: 'A130', format },
          { onReport: () => reports++ },
        ))
          chunks.push(chunk);
      } catch (error) {
        code = (error as any).code;
      }
      failures.push({ format, code, reports, chunks: chunks.length });
    }
    await app.open(id);
    book = app.getWorkbook();
    const target = book.getSheetByName('Export limit')!;
    let preserved = '';
    for await (const chunk of streamExport(book, {
      sheetId: target.id,
      range: 'A130',
      format: 'csv',
      options: { formulas: 'preserve', safe: false },
    }))
      preserved += chunk;
    return { failures, preserved, formula: await target.range('A130').getFormulas() };
  }, formatsModule);
  expect(result.failures).toEqual(
    ['csv', 'tsv', 'html', 'markdown', 'latex'].map((format) => ({
      format,
      code: 'LIMIT_EXCEEDED',
      reports: 0,
      chunks: format === 'html' ? 1 : 0,
    })),
  );
  expect(result.preserved).toBe('=A129+1');
  expect(result.formula).toEqual([['A129+1']]);
});
