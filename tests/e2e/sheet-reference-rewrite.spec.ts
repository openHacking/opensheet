import { expect, test } from '@playwright/test';

test('structural edits preserve case-insensitive sheet references and prevent rebinding', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const values = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const source = await book.addSheet('Limits');
    const target = book.getSheets()[0];
    const other = await book.addSheet('Other');
    await source.range('A1').setValues([[42]]);
    await target.range('A2').setValues([[1]]);
    await other.range('A2').setValues([[10]]);
    await target.range('A1').setFormulas([['limits!A1+A2']]);
    await other.range('A1').setFormulas([['limits!A1+A2']]);
    const read = async () =>
      Promise.all(
        [target, other].map(async (sheet) => (await sheet.range('A1').getValues())[0][0]),
      );
    const before = await read();
    await source.insertRows(0, 1);
    const inserted = await read();
    await book.renameSheet(source.id, 'New Limits');
    const renamed = await read();
    await book.removeSheet(source.id);
    const removed = await read();
    const replacement = await book.addSheet('New Limits');
    await replacement.range('A2').setValues([[123]]);
    const replaced = await read();
    return { before, inserted, renamed, removed, replaced };
  });
  expect(values).toEqual({
    before: [43, 52],
    inserted: [43, 52],
    renamed: [43, 52],
    removed: [{ error: '#REF!' }, { error: '#REF!' }],
    replaced: [{ error: '#REF!' }, { error: '#REF!' }],
  });
});
