import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';
import { parseDelimited } from '@opensheetjs/formats';
import { fileURLToPath } from 'node:url';

test('streams explicitly requested formulas while leaving default exports as values', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const formatModule = `/@fs${fileURLToPath(new URL('../../packages/formats/src/index.ts', import.meta.url))}`;
  const result = await page.evaluate(
    async ({ file, formatModule }) => {
      await window.opensheet.load(JSON.parse(file));
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      await sheet.range('B1').setFormulas([['SUM(1,2)']]);
      await sheet.range('B130').setFormulas([['1/0']]);
      await sheet.range('C130').setValues([['=literal']]);
      const { streamExport } = await import(/* @vite-ignore */ formatModule);
      const chunks: string[] = [];
      for await (const chunk of streamExport(book, {
        sheetId: sheet.id,
        range: 'B1:DX130',
        format: 'csv',
        options: { safe: false, formulas: 'preserve' },
      }))
        chunks.push(chunk);
      const normal: string[] = [];
      for await (const chunk of streamExport(book, {
        sheetId: sheet.id,
        range: 'B1:DX130',
        format: 'csv',
      }))
        normal.push(chunk);
      return { preserved: chunks.join(''), normal: normal.join('') };
    },
    {
      file: JSON.stringify(
        createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] }),
      ),
      formatModule,
    },
  );
  const preserved = parseDelimited(result.preserved),
    normal = parseDelimited(result.normal);
  expect(preserved[0][0]).toBe('=SUM(1,2)');
  expect(preserved[129].slice(0, 2)).toEqual(['=1/0', '=literal']);
  expect(normal[0][0]).toBe('3');
  expect(normal[129].slice(0, 2)).toEqual(['#DIV/0!', "'=literal"]);
});

test('download report bounds address examples and reports zero changes', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(
    async (file) => {
      await window.opensheet.load(JSON.parse(file));
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1:G1')
        .setValues([Array(7).fill('=text')]);
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Data', rows: 1, columns: 7 }] })),
  );
  await page.getByRole('tab', { name: 'Code export' }).click();
  await page.getByRole('tab', { name: 'CSV', exact: true }).click();
  let waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  await waiting;
  await expect(page.locator('#notice')).toContainText(
    'Downloaded CSV: 7 cells prefixed (A1, B1, C1, D1, E1, …)',
  );
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1:G1')
      .setValues([Array(7).fill('plain')]);
  });
  waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  await waiting;
  await expect(page.locator('#notice')).toHaveText('Downloaded CSV: 0 cells prefixed.');
});

test('downloaded CSV applies the safe-text policy to full-width and control prefixes', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(
    async (file) => {
      await window.opensheet.load(JSON.parse(file));
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1:F1')
        .setValues([['＝1', ' ＋1', '\ntext', '\0text', -5, 'ordinary,"=1']]);
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Data', rows: 1, columns: 6 }] })),
  );
  await page.getByRole('tab', { name: 'Code export' }).click();
  await page.getByRole('tab', { name: 'TSV', exact: true }).click();
  await expect(page.locator('#export-note')).toContainText('prefixed with an apostrophe');
  await page.getByRole('tab', { name: 'HTML', exact: true }).click();
  await expect(page.locator('#export-note')).toHaveText('Ready to paste into your next project.');
  await page.getByRole('tab', { name: 'CSV', exact: true }).click();
  await expect(page.locator('#export-note')).toContainText('prefixed with an apostrophe');
  await expect(page.locator('#export-note')).toContainText('imports may change values');
  await expect(page.locator('#export-note')).toContainText(
    'Preview: 4 cells prefixed (A1, B1, C1, D1)',
  );
  const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  await expect(page.locator('#notice')).toContainText(
    'Downloaded CSV: 4 cells prefixed (A1, B1, C1, D1)',
  );
  const stream = await (await waiting).createReadStream();
  const parts: Buffer[] = [];
  for await (const part of stream) parts.push(Buffer.from(part));
  expect(parseDelimited(Buffer.concat(parts).toString('utf8'))).toEqual([
    ["'＝1", "' ＋1", "'\ntext", "'\0text", '-5', 'ordinary,"=1'],
  ]);
  await page.getByRole('tab', { name: 'TSV', exact: true }).click();
  const tsvWaiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  await expect(page.locator('#notice')).toContainText(
    'Downloaded TSV: 4 cells prefixed (A1, B1, C1, D1)',
  );
  const tsvStream = await (await tsvWaiting).createReadStream();
  const tsvParts: Buffer[] = [];
  for await (const part of tsvStream) tsvParts.push(Buffer.from(part));
  expect(parseDelimited(Buffer.concat(tsvParts).toString('utf8'), '\t')).toEqual([
    ["'＝1", "' ＋1", "'\ntext", "'\0text", '-5', 'ordinary,"=1'],
  ]);
});

test('downloaded CSV preserves rows across streaming batches', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(
    async (file) => {
      await window.opensheet.load(JSON.parse(file));
      const sheet = window.opensheet.getWorkbook().getSheets()[0];
      await sheet.range('A128:A130').setValues([['before'], ['=literal'], ['after']]);
      await sheet.range('DX130').setValues([['＠edge']]);
    },
    JSON.stringify(createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] })),
  );
  await page.getByRole('tab', { name: 'Code export' }).click();
  await page.getByRole('tab', { name: 'CSV', exact: true }).click();
  await expect(page.locator('#export-note')).toContainText('Preview is limited');
  await expect(page.locator('#export-note')).toContainText('Preview: 1 cell prefixed (A129)');
  const waiting = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  await expect(page.locator('#notice')).toContainText(
    'Downloaded CSV: 2 cells prefixed (A129, DX130)',
  );
  const download = await waiting;
  const stream = await download.createReadStream();
  const parts: Buffer[] = [];
  for await (const part of stream) parts.push(Buffer.from(part));
  const csv = Buffer.concat(parts).toString('utf8');
  const rows = csv.split('\r\n');
  expect(rows).toHaveLength(130);
  expect(rows[127].split(',')[0]).toBe('before');
  expect(rows[128].split(',')[0]).toBe("'=literal");
  expect(rows[129].split(',')[0]).toBe('after');
  expect(rows[129].split(',')[127]).toBe("'＠edge");
  expect(csv).not.toContain('\uFEFF');
});
