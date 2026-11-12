import { expect, test } from '@playwright/test';
import { createWorkbookFile } from '@opensheetjs/core';

test('imports and reopens maximum merge metadata and rejects overlaps without replacement', async ({
  page,
}) => {
  test.setTimeout(60_000);
  const file = createWorkbookFile({ sheets: [{ name: 'Data', rows: 10000, columns: 2 }] });
  file.sheets[0].merges = Array.from({ length: 10000 }, (_, row) => ({
    startRow: row,
    endRow: row + 1,
    startColumn: 0,
    endColumn: 2,
  })).reverse();
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (fileJSON) => {
    const file = JSON.parse(fileJSON);
    const app = window.opensheet;
    await app.load(file);
    let book = app.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[42]]);
    const id = book.id,
      revision = book.revision;
    const invalid = structuredClone(file);
    invalid.sheets[0].merges[9999] = invalid.sheets[0].merges[0];
    let rejected = false;
    try {
      await app.load(invalid);
    } catch {
      rejected = true;
    }
    const retained = app.getWorkbook().id === id && app.getWorkbook().revision === revision;
    await app.open(id);
    book = app.getWorkbook();
    const exported = await book.toJSON();
    const merges = exported.sheets[0].merges;
    const allValid =
      merges.length === 10000 &&
      merges.every(
        (m, index) =>
          m.startRow === 9999 - index &&
          m.endRow === 10000 - index &&
          m.startColumn === 0 &&
          m.endColumn === 2,
      );
    return {
      rejected,
      retained,
      revision: book.revision,
      expectedRevision: revision,
      allValid,
      value: await book.getSheets()[0].range('A1').getValues(),
    };
  }, JSON.stringify(file));
  expect(result).toEqual({
    rejected: true,
    retained: true,
    revision: result.expectedRevision,
    expectedRevision: result.expectedRevision,
    allValid: true,
    value: [[42]],
  });
});
