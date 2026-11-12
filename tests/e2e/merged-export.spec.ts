import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';
import { fileURLToPath } from 'node:url';

test('exports many merged cells consistently across Worker batches', async ({ page }) => {
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
  file.sheets[0].merges = [
    ...Array.from({ length: 130 }, (_, row) => ({
      startRow: row,
      endRow: row + 1,
      startColumn: 0,
      endColumn: 2,
    })),
    { startRow: 127, endRow: 129, startColumn: 2, endColumn: 3 },
  ].reverse();
  const formatModule = `/@fs${fileURLToPath(new URL('../../packages/formats/src/index.ts', import.meta.url))}`;
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ file, formatModule }) => {
      await window.opensheet.load(JSON.parse(file));
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      await sheet.range('A128').setValues([['merged & text']]);
      await sheet.range('C128').setValues([['vertical']]);
      const { exportRange, streamExport } = await import(/* @vite-ignore */ formatModule);
      const snapshot = await book.toJSON();
      const exports: Record<string, string> = {};
      for (const format of ['html', 'latex'] as const) {
        const config = { sheetId: sheet.id, range: 'A1:DX130', format };
        let streamed = '';
        for await (const chunk of streamExport(book, config)) streamed += chunk;
        if (streamed !== exportRange(snapshot, config).text)
          throw new Error(`${format} stream differs from snapshot`);
        exports[format] = streamed;
      }
      const doc = new DOMParser().parseFromString(exports.html, 'text/html');
      const rows = [...doc.querySelectorAll('tr')];
      return {
        rows: rows.length,
        cells: doc.querySelectorAll('th,td').length,
        horizontal: doc.querySelectorAll('[colspan="2"]').length,
        vertical: doc.querySelectorAll('[rowspan="2"]').length,
        anchor: rows[127].querySelector('td')?.textContent,
        verticalText: rows[127].querySelector('[rowspan="2"]')?.textContent,
        texHorizontal: (exports.latex.match(/\\multicolumn\{2\}/g) ?? []).length,
        texVertical: (exports.latex.match(/\\multirow\{2\}/g) ?? []).length,
        escaped: exports.html.includes('merged &amp; text'),
      };
    },
    { file: JSON.stringify(file), formatModule },
  );
  expect(result).toEqual({
    rows: 130,
    cells: 16509,
    horizontal: 130,
    vertical: 1,
    anchor: 'merged & text',
    verticalText: 'vertical',
    texHorizontal: 130,
    texVertical: 1,
    escaped: true,
  });
});
