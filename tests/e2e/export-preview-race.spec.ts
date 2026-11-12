import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('refresh invalidates the previous preview before the new format can be copied', async ({
  page,
}) => {
  await page.goto('/#code');
  await expect(page.locator('#source')).toContainText('\\begin{tabular}');
  const result = await page.evaluate(() => {
    (document.querySelector('[data-format="csv"]') as HTMLButtonElement).click();
    return {
      disabled: (document.querySelector('#copy-code') as HTMLButtonElement).disabled,
      text: document.querySelector('#source')?.textContent,
      busy: document.querySelector('#format-preview')?.getAttribute('aria-busy'),
    };
  });
  expect(result).toEqual({ disabled: true, text: 'Updating preview…', busy: 'true' });
  await expect(page.locator('#source')).toContainText('Baseline');
  await expect(page.locator('#source')).not.toContainText('\\begin{tabular}');
  await expect(page.getByRole('button', { name: 'Copy preview' })).toBeEnabled();
  await expect(page.locator('#format-preview')).not.toHaveAttribute('aria-busy');
});

test('bounded previews disclose that download contains the full result', async ({ page }) => {
  await page.goto('/#code');
  await expect(page.locator('#source')).toContainText('Baseline');
  await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.insertRows(1, 250);
    await sheet.range('A250').setValues([['outside preview']]);
  });
  await expect(page.locator('#export-note')).toContainText('Preview is limited');
  await expect(page.locator('#source')).not.toContainText('outside preview');
  const download = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  const result = await download;
  const stream = await result.createReadStream();
  let contents = '';
  for await (const chunk of stream) contents += chunk.toString();
  expect(contents).toContain('outside preview');
});

test('an incomplete later export batch never enables copying or publishes a partial download', async ({
  page,
}) => {
  await page.goto('/#code');
  await expect(page.locator('#source')).toContainText('Baseline');
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 130, columns: 128 }] });
  await page.evaluate(async (fileJSON) => {
    const app = window.opensheet;
    await app.load(JSON.parse(fileJSON));
    const book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['first']]);
    await sheet.range('DX130').setValues([['last']]);
    await book.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: sheet.id,
        cells: [{ row: 129, column: 0, input: { type: 'error', code: '#LIMIT!' } }],
      },
    });
    const originalRead = book.readRange.bind(book);
    book.readRange = async (...args) => {
      if (
        args[1].startRow === 128 &&
        args[1].endRow === 130 &&
        args[1].endColumn - args[1].startColumn === 128
      )
        document.body.dataset.lateBatchRead = 'true';
      return originalRead(...args);
    };
    const click = HTMLAnchorElement.prototype.click;
    HTMLAnchorElement.prototype.click = function () {
      document.body.dataset.downloadClicked = 'true';
      return click.call(this);
    };
  }, JSON.stringify(file));
  await page.getByRole('tab', { name: 'CSV', exact: true }).click();
  await expect(page.locator('#source')).toContainText('Cannot export an incomplete calculation');
  await expect(page.locator('#export-note')).toHaveText('Preview unavailable.');
  await expect(page.getByRole('button', { name: 'Copy preview' })).toBeDisabled();
  await page.getByRole('button', { name: 'Download', exact: false }).last().click();
  await expect(page.locator('#notice')).toContainText('Cannot export an incomplete calculation');
  await expect(page.locator('body')).toHaveAttribute('data-late-batch-read', 'true');
  await expect(page.locator('body')).not.toHaveAttribute('data-download-clicked');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A130')
      .setValues([['#LIMIT!']]);
  });
  await expect(page.getByRole('button', { name: 'Copy preview' })).toBeEnabled();
  await expect(page.locator('#source')).toContainText('#LIMIT!');
});
