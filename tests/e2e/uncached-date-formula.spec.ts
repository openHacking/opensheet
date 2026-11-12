import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';

test('rejects an XLSX worksheet that the parser silently drops for an uncached date formula', async ({
  page,
}) => {
  const source = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    source,
    { '!ref': 'A1', A1: { t: 'n', f: 'DATE(2024,1,2)' } },
    'Dates',
  );
  const archive = XLSX.CFB.read(XLSX.write(source, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
  });
  const index = archive.FullPaths.findIndex((path: string) =>
    path.endsWith('worksheets/sheet1.xml'),
  );
  expect(index).toBeGreaterThanOrEqual(0);
  const xml = Buffer.from(archive.FileIndex[index].content).toString('utf8');
  expect(xml).toContain('r="A1" t="e"');
  archive.FileIndex[index].content = Buffer.from(xml.replace('r="A1" t="e"', 'r="A1" t="d"'));
  const buffer = XLSX.CFB.write(archive, { type: 'buffer', fileType: 'zip' });
  const parsed = XLSX.read(buffer, {
    type: 'buffer',
    cellNF: true,
    dense: true,
    cellFormula: true,
    sheetStubs: true,
    cellStyles: true,
    bookVBA: true,
    bookFiles: true,
  });
  expect(parsed.SheetNames).toEqual(['Dates']);
  expect(parsed.Sheets.Dates).toBeUndefined();
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('A1').setValues([[8]]);
    return { id: book.id, revision: book.revision, bytes: (await book.storageStats()).bytes };
  });
  await page.locator('#file').setInputFiles({
    name: 'uncached-date.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer,
  });
  await expect(page.locator('#notice')).toContainText('Missing worksheet "Dates"');
  await expect(page.locator('#notice')).toContainText('the parser could not read it');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    const unchanged = {
      id: book.id,
      revision: book.revision,
      bytes: (await book.storageStats()).bytes,
      value: await sheet.range('A1').getValues(),
    };
    const undo = await book.undo(),
      undone = await sheet.range('A1').getValues();
    const redo = await book.redo(),
      redone = await sheet.range('A1').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { unchanged, undo, undone, redo, redone, reopened: await sheet.range('A1').getValues() };
  });
  expect(result.unchanged).toEqual({ ...before, value: [[8]] });
  expect(result.undo).toBe(true);
  expect(result.undone).toEqual([[7]]);
  expect(result.redo).toBe(true);
  expect(result.redone).toEqual([[8]]);
  expect(result.reopened).toEqual([[8]]);
});
