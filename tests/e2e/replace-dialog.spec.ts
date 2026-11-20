import { workbookTool } from './workbook-tools.js';
import { expect, test } from '@playwright/test';

test.beforeEach(async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1:A2')
      .setValues([['Alpha'], ['alpha']]);
  });
});

test('cancel and validation preserve data, then one replacement is undoable and persistent', async ({
  page,
}) => {
  await workbookTool(page, 'replace');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(page.getByLabel('Find text')).toBeFocused();
  await page.getByLabel('Find text').fill('Alpha');
  await page.getByLabel('Replace with').fill('Omega');
  await page.locator('#dialog-close').click();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    ),
  ).toEqual([['Alpha'], ['alpha']]);
  await workbookTool(page, 'replace');
  await page.getByLabel('Find text').fill('Alpha');
  await page.getByLabel('Replace with').fill('Omega');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  const values = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    const replaced = await sheet.range('A1:A2').getValues();
    await book.undo();
    const undone = await sheet.range('A1:A2').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    return { replaced, undone, reopened: await book.getSheets()[0].range('A1:A2').getValues() };
  });
  expect(values).toEqual({
    replaced: [['Omega'], ['alpha']],
    undone: [['Alpha'], ['alpha']],
    reopened: [['Omega'], ['alpha']],
  });
});

test('keyboard users can identify, navigate and dismiss the dialog', async ({ page }) => {
  const actions = page.getByRole('button', { name: 'Replace literal text' });
  await actions.focus();
  await actions.press('Enter');
  await expect(page.getByRole('dialog', { name: 'Replace text' })).toBeVisible();
  await expect(page.getByLabel('Find text')).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(page.getByRole('button', { name: 'Close dialog' })).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByLabel('Find text')).toBeFocused();
  await page.getByLabel('Find text').fill('Alpha');
  await page.keyboard.press('Tab');
  await expect(page.getByLabel('Replace with')).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(page.getByRole('button', { name: 'Replace all' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog', { name: 'Replace text' })).not.toBeVisible();
  await expect(actions).toBeFocused();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    ),
  ).toEqual([['Alpha'], ['alpha']]);
});

test('stale revision and read-only mode cannot replace text', async ({ page }) => {
  await workbookTool(page, 'replace');
  await page.getByLabel('Find text').fill('Alpha');
  await page.getByLabel('Replace with').fill('Omega');
  const revision = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A2')
      .setValues([['new edit']]);
    return book.revision;
  });
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(page.getByRole('alert')).toHaveText(
    'Workbook changed. Close this dialog and try again.',
  );
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(revision);
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    ),
  ).toEqual([['Alpha'], ['new edit']]);
  await page.locator('#dialog-close').click();
  await page.getByLabel('Read only', { exact: true }).check();
  await workbookTool(page, 'replace');
  await expect(page.getByRole('button', { name: 'Replace all' })).toBeDisabled();
  await expect(page.getByRole('alert')).toHaveText('Read only. Replacement is disabled.');
});

test('replacement metacharacters remain literal text', async ({ page }) => {
  await workbookTool(page, 'replace');
  await page.getByLabel('Find text').fill('Alpha');
  await page.getByLabel('Replace with').fill('$&');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    ),
  ).toEqual([['$&'], ['alpha']]);
  await page.evaluate(() => window.opensheet.getWorkbook().undo());
  expect(
    await page.evaluate(() =>
      window.opensheet.getWorkbook().getSheets()[0].range('A1:A2').getValues(),
    ),
  ).toEqual([['Alpha'], ['alpha']]);
});

test('replacement with no changes preserves redo across reopen', async ({ page }) => {
  const before = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A3')
      .setValues([['redo value']]);
    await book.undo();
    return { id: book.id, revision: book.revision, canRedo: book.canRedo };
  });
  expect(before.canRedo).toBe(true);
  await workbookTool(page, 'replace');
  await page.getByLabel('Find text').fill('not-present');
  await page.getByLabel('Replace with').fill('new');
  await page.getByRole('button', { name: 'Replace all' }).click();
  await expect(page.getByRole('dialog')).not.toBeVisible();
  await expect(page.locator('#notice')).toContainText('No text changed');
  const after = await page.evaluate(async (id) => {
    const app = window.opensheet;
    const revision = app.getWorkbook().revision;
    await app.open(id);
    const book = app.getWorkbook();
    const canRedo = book.canRedo;
    const redo = await book.redo();
    return { revision, canRedo, redo, value: await book.getSheets()[0].range('A3').getValues() };
  }, before.id);
  expect(after).toEqual({
    revision: before.revision,
    canRedo: true,
    redo: true,
    value: [['redo value']],
  });
});

test('replacement accepts the exact text limit and rejects one unit above it', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([['a'.repeat(16384)]]);
    await book.replaceText(sheet.id, 'a', 'aa');
    const exactLength = ((await sheet.range('A1').getValues())[0][0] as string).length;
    await book.undo();
    await sheet.range('A1').setValues([['a'.repeat(16385)]]);
    const revision = book.revision;
    let error = '';
    try {
      await book.replaceText(sheet.id, 'a', 'aa');
    } catch (reason) {
      error = (reason as { code?: string }).code ?? '';
    }
    const afterRevision = book.revision;
    const id = book.id;
    await app.open(id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const lengthAfterReopen = ((await sheet.range('A1').getValues())[0][0] as string).length;
    await book.undo();
    const lengthAfterUndo = ((await sheet.range('A1').getValues())[0][0] as string).length;
    return { exactLength, error, revision, afterRevision, lengthAfterReopen, lengthAfterUndo };
  });
  expect(result).toEqual({
    exactLength: 32768,
    error: 'LIMIT_EXCEEDED',
    revision: result.revision,
    afterRevision: result.revision,
    lengthAfterReopen: 16385,
    lengthAfterUndo: 16384,
  });
});
