import { expect, test } from '@playwright/test';

test('rejected annotation across a merge leaves the revision and history unchanged', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:C3').clear({ all: true });
    await sheet.range('B2:C3').merge();
    await sheet.range('A1').setNote('original');
    const revision = book.revision;
    let error = '';
    try {
      await sheet.range('A1:C3').setNote('changed');
    } catch (cause) {
      error = (cause as { code?: string }).code ?? '';
    }
    const after = { revision: book.revision, notes: await sheet.range('A1:C3').getNotes() };
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1:C3').getNotes();
    await book.undo();
    const undone = await sheet.range('A1').getNotes();
    await book.redo();
    return { revision, error, after, reopened, undone, redone: await sheet.range('A1').getNotes() };
  });
  expect(result.error).toBe('INVALID_RANGE');
  expect(result.after.revision).toBe(result.revision);
  expect(result.after.notes).toEqual([
    ['original', null, null],
    [null, null, null],
    [null, null, null],
  ]);
  expect(result.reopened).toEqual(result.after.notes);
  expect(result.undone).toEqual([[null]]);
  expect(result.redone).toEqual([['original']]);
});
