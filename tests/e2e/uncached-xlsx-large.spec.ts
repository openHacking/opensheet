import { expect, test } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';

test('imports every uncached formula and formatted blank through reopen', async ({ page }) => {
  test.setTimeout(120000);
  const bytes = readFileSync(
    new URL('../../benchmarks/uncached-import-10k-2026-11-11.xlsx', import.meta.url),
  );
  const fixture = JSON.parse(
    readFileSync(
      new URL('../../benchmarks/uncached-import-10k-2026-11-11.xlsx.json', import.meta.url),
      'utf8',
    ),
  );
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(fixture.sha256);
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.locator('#file').setInputFiles({
    name: 'uncached-large.xlsx',
    mimeType: 'application/octet-stream',
    buffer: bytes,
  });
  const dialog = page.getByRole('dialog', { name: 'Compatibility report' });
  await expect(dialog).toBeVisible();
  const confirm = dialog.getByRole('button', { name: 'Import converted workbook' });
  if (await confirm.isVisible()) await confirm.click();
  await expect(page.getByRole('tab', { name: 'Uncached benchmark', exact: true })).toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const verify = async () => {
      const formulas = await sheet.range('B1:B10000').getFormulas();
      const values = await sheet.range('B1:B10000').getValues();
      const blanks = await book.readCells(sheet.id, sheet.range('C1:C10000').bounds, {
        fields: ['input', 'numberFormat'],
      });
      let verified = 0;
      for (let row = 0; row < 10000; row++) {
        if (formulas[row][0] !== `A${row + 1}*2` || values[row][0] !== (row + 1) * 2)
          throw new Error(`Formula mismatch at row ${row + 1}`);
        const blank = blanks.cells[row];
        if (blank.row !== row || blank.input?.type !== 'blank')
          throw new Error(`Blank mismatch at row ${row + 1}`);
        if (blank.numberFormat !== (row % 5 === 4 ? '0.00' : undefined))
          throw new Error(`Number format mismatch at row ${row + 1}`);
        verified++;
      }
      return verified;
    };
    const initial = await verify();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    return { initial, reopened: await verify() };
  });
  expect(result).toEqual({ initial: 10000, reopened: 10000 });
});
