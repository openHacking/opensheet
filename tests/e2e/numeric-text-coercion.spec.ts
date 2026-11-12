import { expect, test } from '@playwright/test';

test('radix-prefixed text remains invalid through recalculation, history and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = await book.addSheet('Numeric text', { rows: 4, columns: 3 });
    await sheet.range('A1:A4').setValues([['0x10'], [' 0B10 '], ['0o10'], ['1e2']]);
    await sheet
      .range('B1:B4')
      .setFormulas([['A1+1'], ['ROUND(A2,0)'], ['ABS(A3)'], ['ROUND(A4,0)']]);
    await sheet
      .range('C1:C4')
      .setFormulas([
        ['MID("abc","0x1",1)'],
        ['DATE("0x7e8",1,1)'],
        ['INDEX(A1:A4,"0b10")'],
        ['MATCH(2,A1:A4,"0o0")'],
      ]);
    const initial = await sheet.range('B1:B4').getValues();
    await sheet.range('A1').setValues([['16']]);
    const changed = await sheet.range('B1:B4').getValues();
    await book.undo();
    const undone = await sheet.range('B1:B4').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheetByName('Numeric text')!;
    return {
      initial,
      changed,
      undone,
      reopened: await sheet.range('B1:B4').getValues(),
      contexts: await sheet.range('C1:C4').getValues(),
    };
  });
  const invalid = { error: '#VALUE!' };
  expect(result.initial).toEqual([[invalid], [invalid], [invalid], [100]]);
  expect(result.undone).toEqual(result.initial);
  expect(result.changed).toEqual([[17], [invalid], [invalid], [100]]);
  expect(result.reopened).toEqual(result.changed);
  expect(result.contexts).toEqual([[invalid], [invalid], [invalid], [invalid]]);
});
