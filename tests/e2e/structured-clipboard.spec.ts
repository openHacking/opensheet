import { workbookTool } from './workbook-tools.js';
import { test, expect } from '@playwright/test';
import { readForPoll } from './read-for-poll.js';

test('copy event publishes HTML and text through ClipboardItem promises without changing cells', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const initial = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[42]]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async (items: ClipboardItem[]) => {
          document.body.dataset.copiedText = await (await items[0].getType('text/plain')).text();
          document.body.dataset.copiedHtml = await (await items[0].getType('text/html')).text();
        },
        writeText: async (text: string) => {
          document.body.dataset.copiedText = text;
        },
      },
    });
    return { revision: book.revision, rich: typeof ClipboardItem !== 'undefined' };
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('copy');
  await expect(page.locator('body')).toHaveAttribute('data-copied-text', '42');
  if (initial.rich)
    await expect(page.locator('body')).toHaveAttribute(
      'data-copied-html',
      /data-opensheet-clipboard="1"/,
    );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(initial.revision);
  await page.evaluate(() => {
    Object.defineProperty(window, 'ClipboardItem', { configurable: true, value: undefined });
    delete document.body.dataset.copiedText;
    delete document.body.dataset.copiedHtml;
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('copy');
  await expect(page.locator('body')).toHaveAttribute('data-copied-text', '42');
  await expect(page.locator('body')).not.toHaveAttribute('data-copied-html');
});

test('structured copy preserves exact inputs, relocates formulas and pastes captured values', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const data = await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:F4').clear({ all: true });
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('A1').setNote('copied <script>literal</script>');
    await sheet.range('A1').setLink({ target: 'https://example.com/source', tooltip: 'Source' });
    await sheet.range('C2:D2').setNote('old target');
    await sheet.range('B1').setFormulas([['A1+$A$1']]);
    await sheet.range('A1:B1').setStyle({ bold: true, numberFormat: '0.00' });
    await sheet.range('C1').setValues([['<script>literal</script>\r\n"quoted"']]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 3 });
    return app.getGrid()!.copyClipboard();
  });
  expect(data.text).toContain('7.00\t14.00');
  expect(data.html).not.toContain('<script>literal</script>');
  await page.evaluate((data) => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 2, endColumn: 3 });
    app.getGrid()!.pasteHTML(data.html, data.text);
  }, data);
  await expect
    .poll(() => readForPoll(page, 'C2:E2'))
    .toEqual([[7, 14, '<script>literal</script>\r\n"quoted"']]);
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      return {
        formulas: await sheet.range('C2:D2').getFormulas(),
        display: await sheet.range('C2:D2').getDisplayValues(),
      };
    }),
  ).toEqual({ formulas: [[null, '(C2+$A$1)']], display: [['7.00', '14.00']] });
  expect(
    await page.evaluate(async () => {
      const sheet = window.opensheet.getWorkbook().getSheets()[0];
      return {
        notes: await sheet.range('C2:D2').getNotes(),
        links: await sheet.range('C2:D2').getLinks(),
      };
    }),
  ).toEqual({
    notes: [['copied <script>literal</script>', null]],
    links: [[{ target: 'https://example.com/source', tooltip: 'Source' }, null]],
  });
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1').setValues([[100]]);
    await sheet.range('A4').setNote('keep target');
    app.selection.set({ sheetId: sheet.id, startRow: 3, endRow: 4, startColumn: 0, endColumn: 1 });
  });
  await workbookTool(page, 'paste-special');
  await page.getByLabel('Paste clipboard here').evaluate((element, data) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', data.text);
    clipboardData.setData('text/html', data.html);
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    element.dispatchEvent(event);
  }, data);
  await page.getByLabel('Paste mode').selectOption('values');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  await expect
    .poll(() => readForPoll(page, 'A4:C4'))
    .toEqual([[7, 14, '<script>literal</script>\r\n"quoted"']]);
  expect(
    await page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].range('A4').getNotes()),
  ).toEqual([['keep target']]);
});

test('structured values paste keeps exact strings and captured calculated values', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const data = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1:C2').clear({ all: true });
    await sheet.range('A1:B1').setValues([["'", "'=1+1"]]);
    await sheet.range('C1').setFormulas([['1+1']]);
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 3 });
    const copied = await app.getGrid()!.copyClipboard();
    await sheet.range('C1').setFormulas([['2+2']]);
    app.selection.set({ sheetId: sheet.id, startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 });
    return copied;
  });
  await workbookTool(page, 'paste-special');
  await page.getByLabel('Paste clipboard here').evaluate((element, payload) => {
    const clipboardData = new DataTransfer();
    clipboardData.setData('text/plain', payload.text);
    clipboardData.setData('text/html', payload.html);
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    element.dispatchEvent(event);
  }, data);
  await page.getByLabel('Paste mode').selectOption('values');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A2:C2').getValues(),
    ),
  ).toEqual([["'", "'=1+1", 2]]);
});

test('malformed structured payloads fail atomically', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const original = await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    return { revision: app.getWorkbook().revision, values: await sheet.range('A1').getValues() };
  });
  await page.evaluate(() =>
    window.opensheet
      .getGrid()!
      .pasteHTML(
        '<table data-opensheet-clipboard="1" data-row="0" data-column="0"><tr><td data-opensheet-cell="{}">99</td></tr></table>',
      ),
  );
  await expect(page.locator('#notice')).toContainText('Invalid structured clipboard cell');
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return { revision: book.revision, values: await book.getSheets()[0].range('A1').getValues() };
    }),
  ).toEqual(original);
});

test('rejects unknown structured fields atomically and preserves legacy omitted annotations', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const original = await page.evaluate(async () => {
    const app = window.opensheet;
    const sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1').setValues([[123]]);
    await sheet.range('A1').setNote('keep');
    await sheet.range('A1').setLink({ target: 'https://example.com/' });
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
    return app.getWorkbook().revision;
  });
  const payload = {
    input: { type: 'number', value: 7 },
    value: { type: 'number', value: 7 },
    style: {},
  };
  for (const bad of [
    { ...payload, futureSemantics: true },
    { ...payload, input: { ...payload.input, futureSemantics: true } },
    { ...payload, value: { ...payload.value, futureSemantics: true } },
    { ...payload, input: { type: 'formula', expression: '1+1', futureSemantics: true } },
  ]) {
    await page.evaluate((value) => {
      document.querySelector('#notice')!.textContent = '';
      const table = document.createElement('table');
      table.dataset.opensheetClipboard = '1';
      table.dataset.row = '0';
      table.dataset.column = '0';
      const row = table.insertRow();
      row.insertCell().dataset.opensheetCell = JSON.stringify({
        input: { type: 'number', value: 8 },
        value: { type: 'number', value: 8 },
        style: {},
      });
      row.insertCell().dataset.opensheetCell = JSON.stringify(value);
      window.opensheet.getGrid()!.pasteHTML(table.outerHTML);
    }, bad);
    await expect(page.locator('#notice')).toContainText('Invalid structured clipboard cell');
    expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(original);
    expect(
      await page.evaluate(() =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[123]]);
  }
  await page.evaluate((value) => {
    const table = document.createElement('table');
    table.dataset.opensheetClipboard = '1';
    table.dataset.row = '0';
    table.dataset.column = '0';
    table.insertRow().insertCell().dataset.opensheetCell = JSON.stringify(value);
    window.opensheet.getGrid()!.pasteHTML(table.outerHTML);
  }, payload);
  await expect
    .poll(() => page.evaluate(() => window.opensheet.getWorkbook().revision))
    .toBe(original + 1);
  const restored = await page.evaluate(async () => {
    const sheet = window.opensheet.getWorkbook().getSheets()[0];
    return {
      value: await sheet.range('A1').getValues(),
      note: await sheet.range('A1').getNotes(),
      link: await sheet.range('A1').getLinks(),
    };
  });
  expect(restored).toEqual({
    value: [[7]],
    note: [['keep']],
    link: [[{ target: 'https://example.com/' }]],
  });
});
