import { expect, test } from '@playwright/test';

test('reference overflow remains parseable through copy, structural edits and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = await book.addSheet('Reference bounds', { rows: 5, columns: 4 });
    await sheet.range('A1:B1').setFormulas([['IFERROR(A1048576,42)', 'IFERROR(XFD1,42)']]);
    await sheet.range('A1').copyTo('A2');
    await sheet.range('B1').copyTo('C1');
    const copies = [await sheet.range('A2').getFormulas(), await sheet.range('C1').getFormulas()];
    await sheet.insertRows(0);
    await sheet.insertColumns(0);
    const inserted = await sheet.range('B2:C2').getFormulas();
    const values = await sheet.range('B2:C2').getValues();
    await book.undo();
    await book.undo();
    const undone = await sheet.range('A1:B1').getFormulas();
    await book.redo();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheetByName('Reference bounds')!;
    return { copies, inserted, values, undone, reopened: await sheet.range('B2:C2').getFormulas() };
  });
  for (const formula of [result.copies[0][0][0], result.copies[1][0][0]])
    expect(formula).toContain('#REF!');
  expect(result.inserted).toEqual([['IFERROR(#REF!,42)', 'IFERROR(#REF!,42)']]);
  expect(result.values).toEqual([[42, 42]]);
  expect(result.undone).toEqual([['IFERROR(A1048576,42)', 'IFERROR(XFD1,42)']]);
  expect(result.reopened).toEqual(result.inserted);
});
