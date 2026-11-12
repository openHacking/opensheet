import { test, expect } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  await page.evaluate(() => {
    const app = window.opensheet,
      sheet = app.getWorkbook().getSheets()[0];
    app.selection.set({ sheetId: sheet.id, startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 });
  });
});

test('edits literal notes and links, restores history and presents a safe read-only view', async ({
  page,
}) => {
  const note = '备注\n<script>window.annotationInjected = true</script>';
  await page.getByLabel('Sheet actions').selectOption('annotation');
  await expect(page.getByLabel('Cell note', { exact: true })).toBeEditable();
  await page.getByLabel('Cell note', { exact: true }).fill(note);
  await page.getByLabel('Link URL', { exact: true }).fill('https://example.com/source');
  await page.getByLabel('Link description', { exact: true }).fill('Source document');
  await expect(page.getByRole('link', { name: 'Open link' })).toHaveAttribute(
    'rel',
    'noopener noreferrer',
  );
  await page.getByRole('button', { name: 'Save note and link' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet,
      book = app.getWorkbook(),
      range = book.getSheets()[0].range('A1');
    const saved = { notes: await range.getNotes(), links: await range.getLinks() };
    await book.undo();
    const undone = await range.getNotes();
    await book.redo();
    await app.open(book.id);
    return {
      saved,
      undone,
      restored: await app.getWorkbook().getSheets()[0].range('A1').getNotes(),
    };
  });
  expect(result.saved).toEqual({
    notes: [[note]],
    links: [[{ target: 'https://example.com/source', tooltip: 'Source document' }]],
  });
  expect(result.undone).toEqual([[null]]);
  expect(result.restored).toEqual([[note]]);
  await expect(page.getByRole('gridcell')).toContainText('Has note. Has link');
  await page.getByLabel('Read only', { exact: true }).check();
  await page.getByLabel('Sheet actions').selectOption('annotation');
  await expect(page.getByLabel('Cell note', { exact: true })).toHaveValue(note);
  await expect(page.getByLabel('Cell note', { exact: true })).not.toBeEditable();
  await expect(page.getByRole('button', { name: 'Save note and link' })).toBeDisabled();
  await expect(page.getByRole('link', { name: 'Open link' })).toHaveAttribute(
    'href',
    'https://example.com/source',
  );
  expect(await page.evaluate(() => Object.hasOwn(window, 'annotationInjected'))).toBe(false);
});

test('blocks unsafe URLs and stale changes without publishing annotations', async ({ page }) => {
  await page.getByLabel('Sheet actions').selectOption('annotation');
  await expect(page.getByLabel('Cell note', { exact: true })).toBeEditable();
  const original = await page.evaluate(() => window.opensheet.getWorkbook().revision);
  await page.getByLabel('Link URL', { exact: true }).fill('javascript:alert(1)');
  await expect(page.getByRole('link', { name: 'Open link' })).not.toBeVisible();
  await page.getByRole('button', { name: 'Save note and link' }).click();
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(original);
  expect(
    await page
      .getByLabel('Link URL', { exact: true })
      .evaluate((el: HTMLInputElement) => el.validity.valid),
  ).toBe(false);
  await page.getByLabel('Link URL', { exact: true }).fill('https://example.com');
  const changed = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A2')
      .setValues([[42]]);
    return book.revision;
  });
  await page.getByRole('button', { name: 'Save note and link' }).click();
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Close this dialog and select the cell again.',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(changed);
});

test('removes annotations together and restores them with undo', async ({ page }) => {
  await page.evaluate(async () => {
    const range = window.opensheet.getWorkbook().getSheets()[0].range('A1');
    await range.setNote('old');
    await range.setLink({ target: 'mailto:hello@example.com' });
  });
  await page.getByLabel('Sheet actions').selectOption('annotation');
  await expect(page.getByLabel('Cell note', { exact: true })).toHaveValue('old');
  await page.getByRole('button', { name: 'Remove note and link' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook(),
      range = book.getSheets()[0].range('A1');
    const removed = [await range.getNotes(), await range.getLinks()];
    await book.undo();
    return { removed, restored: [await range.getNotes(), await range.getLinks()] };
  });
  expect(result.removed).toEqual([[[null]], [[null]]]);
  expect(result.restored).toEqual([[['old']], [[{ target: 'mailto:hello@example.com' }]]]);
});

test('empty note on a blank cell survives persistence, reopen and history', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').clear({ all: true });
    await sheet.range('A1').setNote('');
    const saved = await sheet.range('A1').getNotes();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1').getNotes();
    const undo = await book.undo(),
      undone = await sheet.range('A1').getNotes();
    const redo = await book.redo(),
      redone = await sheet.range('A1').getNotes();
    return { saved, reopened, undo, undone, redo, redone };
  });
  expect(result).toEqual({
    saved: [['']],
    reopened: [['']],
    undo: true,
    undone: [[null]],
    redo: true,
    redone: [['']],
  });
});

test('editing a link preserves an untouched empty note and clearing a note removes it', async ({
  page,
}) => {
  await page.evaluate(async () => {
    const range = window.opensheet.getWorkbook().getSheets()[0].range('A1');
    await range.setNote('');
  });
  await page.getByLabel('Sheet actions').selectOption('annotation');
  await expect(page.getByLabel('Cell note', { exact: true })).toBeEditable();
  await page.getByLabel('Link URL', { exact: true }).fill('https://example.com/empty-note');
  await page.getByRole('button', { name: 'Save note and link' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let range = book.getSheets()[0].range('A1');
    const saved = [await range.getNotes(), await range.getLinks()];
    await book.undo();
    const undone = [await range.getNotes(), await range.getLinks()];
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    range = book.getSheets()[0].range('A1');
    const reopened = [await range.getNotes(), await range.getLinks()];
    await range.setNote('erase me');
    return { saved, undone, reopened };
  });
  const link = { target: 'https://example.com/empty-note' };
  expect(result.saved).toEqual([[['']], [[link]]]);
  expect(result.undone).toEqual([[['']], [[null]]]);
  expect(result.reopened).toEqual([[['']], [[link]]]);
  await page.getByLabel('Sheet actions').selectOption('annotation');
  const note = page.getByLabel('Cell note', { exact: true });
  await expect(note).toHaveValue('erase me');
  await note.fill('');
  await page.getByRole('button', { name: 'Save note and link' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(async () => {
      const range = window.opensheet.getWorkbook().getSheets()[0].range('A1');
      return [await range.getNotes(), await range.getLinks()];
    }),
  ).toEqual([[[null]], [[link]]]);
});

test('editing a note preserves an untouched empty link description and supports clearing it', async ({
  page,
}) => {
  const target = 'https://example.com/empty-description';
  await page.evaluate(async (target) => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setLink({ target, tooltip: '' });
  }, target);
  await page.getByLabel('Sheet actions').selectOption('annotation');
  await expect(page.getByLabel('Cell note', { exact: true })).toBeEditable();
  await page.getByLabel('Cell note', { exact: true }).fill('edited');
  await page.getByRole('button', { name: 'Save note and link' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let range = book.getSheets()[0].range('A1');
    const saved = await range.getLinks();
    await book.undo();
    const undone = await range.getLinks();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    range = book.getSheets()[0].range('A1');
    const reopened = await range.getLinks();
    await range.setLink({ target: 'https://example.com/empty-description', tooltip: 'remove me' });
    return { saved, undone, reopened };
  });
  expect(result.saved).toEqual([[{ target, tooltip: '' }]]);
  expect(result.undone).toEqual([[{ target, tooltip: '' }]]);
  expect(result.reopened).toEqual([[{ target, tooltip: '' }]]);
  await page.getByLabel('Sheet actions').selectOption('annotation');
  const tooltip = page.getByLabel('Link description', { exact: true });
  await expect(tooltip).toHaveValue('remove me');
  await tooltip.fill('');
  await page.getByRole('button', { name: 'Save note and link' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(() => window.opensheet.getWorkbook().getSheets()[0].range('A1').getLinks()),
  ).toEqual([[{ target }]]);
});
