import { test, expect, type Page } from '@playwright/test';

async function capture(page: Page, text: string, html = '') {
  await page.getByLabel('Paste clipboard here').evaluate(
    (element, data) => {
      const clipboardData = new DataTransfer();
      clipboardData.setData('text/plain', data.text);
      if (data.html) clipboardData.setData('text/html', data.html);
      const event = new Event('paste', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: clipboardData });
      element.dispatchEvent(event);
    },
    { text, html },
  );
}

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(async () => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    await sheet.range('A1').clear({ all: true });
    await sheet.range('A1').setValues([[7]]);
    await sheet.range('A1').setStyle({ italic: true });
    await sheet.range('A1').setNote('keep');
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  });
  await page.getByLabel('Sheet actions').selectOption('paste-special');
});

test('image-only clipboard cannot clear cells or apply previously captured text', async ({
  page,
}) => {
  await capture(page, '99');
  await page.getByLabel('Paste clipboard here').evaluate((element) => {
    const clipboardData = new DataTransfer();
    clipboardData.items.add(new File(['image'], 'image.png', { type: 'image/png' }));
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    element.dispatchEvent(event);
  });
  await expect(page.getByRole('alert')).toHaveText(
    'Clipboard has no supported text or HTML table.',
  );
  await expect(page.getByRole('button', { name: 'Apply paste' })).toBeDisabled();
  await page.locator('#dialog-close').click();
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByRole('grid', { exact: true }).evaluate((element) => {
    const clipboardData = new DataTransfer();
    clipboardData.items.add(new File(['image'], 'image.png', { type: 'image/png' }));
    const event = new Event('paste', { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'clipboardData', { value: clipboardData });
    element.dispatchEvent(event);
  });
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return { revision: book.revision, values: await book.getSheets()[0].range('A1').getValues() };
    }),
  ).toEqual({ revision, values: [[7]] });
});

test('malformed quoted text fails atomically and a corrected paste remains undoable', async ({
  page,
}) => {
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await capture(page, '"new"tail');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('alert')).toHaveText('Unexpected text after closing quote');
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      return { revision: book.revision, values: await book.getSheets()[0].range('A1').getValues() };
    }),
  ).toEqual({ revision, values: [[7]] });
  await capture(page, '"new ""quote"""');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      const value = await sheet.range('A1').getValues();
      await book.undo();
      return { value, undone: await sheet.range('A1').getValues() };
    }),
  ).toEqual({ value: [['new "quote"']], undone: [[7]] });
});

for (const kind of ['text', 'html'] as const) {
  test(`values-only ${kind} paste preserves external leading apostrophes`, async ({ page }) => {
    const before = await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:F1').getValues(),
    );
    await capture(
      page,
      "'O'Brien\t'=1+1\t'\t=1+1\t001\tTRUE",
      kind === 'html'
        ? "<table><tr><td>'O'Brien</td><td>'=1+1</td><td>'</td><td>=1+1</td><td>001</td><td>TRUE</td></tr></table>"
        : '',
    );
    await page.getByLabel('Paste mode').selectOption('values');
    await page.getByRole('button', { name: 'Apply paste' }).click();
    await expect(page.getByRole('dialog')).not.toBeVisible();
    const result = await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook();
      const sheet = book.getSheets()[0];
      const value = await sheet.range('A1:F1').getValues();
      await book.undo();
      const undone = await sheet.range('A1:F1').getValues();
      await book.redo();
      await window.opensheet.open(book.id);
      return {
        value,
        undone,
        reopened: await window.opensheet.getWorkbook().getSheets()[0].range('A1:F1').getValues(),
      };
    });
    expect(result.value).toEqual([["'O'Brien", "'=1+1", "'", '=1+1', '001', true]]);
    expect(result.undone).toEqual(before);
    expect(result.reopened).toEqual(result.value);
  });
}

test('formula-input paste retains the apostrophe editing prefix', async ({ page }) => {
  await capture(page, "'=1+1\t'\t=1+1");
  await page.getByLabel('Paste mode').selectOption('formulas');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:C1').getValues(),
    ),
  ).toEqual([['=1+1', '', 2]]);
});

test('captures without writing, then applies values literally and preserves destination format', async ({
  page,
}) => {
  const revision = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await capture(page, '=1+2', '<table><tr><td style="font-weight:bold">=1+2</td></tr></table>');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  await page.getByLabel('Paste mode').selectOption('values');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      return {
        value: await sheet.range('A1').getValues(),
        style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
        notes: await sheet.range('A1').getNotes(),
      };
    }),
  ).toEqual({ value: [['=1+2']], style: { italic: true }, notes: [['keep']] });
});

test('formula-input mode evaluates formulas while preserving formatting and supports undo', async ({
  page,
}) => {
  await capture(page, '=1+2');
  await page.getByLabel('Paste mode').selectOption('formulas');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      const value = await sheet.range('A1').getValues(),
        formula = await sheet.range('A1').getFormulas();
      await book.undo();
      return { value, formula, undone: await sheet.range('A1').getValues() };
    }),
  ).toEqual({ value: [[3]], formula: [['1+2']], undone: [[7]] });
});

test('formats require HTML and replace only formatting', async ({ page }) => {
  await capture(page, '99');
  await page.getByLabel('Paste mode').selectOption('formats');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('alert')).toContainText('Formats require an HTML table');
  await capture(page, '99', '<table><tr><td style="font-weight:bold">99</td></tr></table>');
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(async () => {
      const book = window.opensheet.getWorkbook(),
        sheet = book.getSheets()[0];
      return {
        values: await sheet.range('A1').getValues(),
        style: book.getStyle((await book.getCell(sheet.id, 0, 0))?.styleId),
      };
    }),
  ).toEqual({ values: [[7]], style: { bold: true } });
});

test('cancel, stale dialog and read-only mode preserve cells', async ({ page }) => {
  await capture(page, '99');
  await page.locator('#dialog-close').click();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[7]]);
  await page.getByLabel('Sheet actions').selectOption('paste-special');
  await capture(page, '99');
  await page.evaluate(() =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[8]]),
  );
  await page.getByRole('button', { name: 'Apply paste' }).click();
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Close this dialog and select the target again.',
  );
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
    ),
  ).toEqual([[8]]);
  await page.locator('#dialog-close').click();
  await page.getByLabel('Read only', { exact: true }).check();
  await page.getByLabel('Sheet actions').selectOption('paste-special');
  await expect(page.getByRole('alert')).toHaveText('Read only. Pasting is disabled.');
  await expect(page.getByRole('button', { name: 'Apply paste' })).toBeDisabled();
});
