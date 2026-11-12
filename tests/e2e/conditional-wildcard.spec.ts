import { expect, test } from '@playwright/test';

test('conditional comparisons bound long literal and numeric text across history', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    const text = 'x'.repeat(32767),
      numericText = '0'.repeat(32767);
    await sheet.range('A1:B30').setValues(Array.from({ length: 30 }, () => [text, numericText]));
    await sheet.range('C1:C2').setFormulas([['COUNTIF(A1:A31,"needle")'], ['COUNTIF(B1:B31,1)']]);
    const initial = await sheet.range('C1:C2').getValues();
    await sheet.range('A31:B31').setValues([[text, numericText]]);
    const limited = await sheet.range('C1:C2').getValues();
    await book.undo();
    const undone = await sheet.range('C1:C2').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('C1:C2').getValues();
    await sheet.range('A31:B31').setValues([['x', '0']]);
    return {
      initial,
      limited,
      undone,
      reopened,
      recovered: await sheet.range('C1:C2').getValues(),
    };
  });
  const limit = [[{ error: '#LIMIT!' }], [{ error: '#LIMIT!' }]];
  expect(result).toEqual({
    initial: [[0], [0]],
    limited: limit,
    undone: [[0], [0]],
    reopened: limit,
    recovered: [[0], [0]],
  });
});

test('conditional projected values honor sheet bounds through edits and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = await book.addSheet('Projected bounds', { rows: 100, columns: 26 });
    await sheet.range('A1:B2').setValues([
      [1, 1],
      [1, null],
    ]);
    await sheet.range('B100').setValues([[2]]);
    await sheet.range('Z1').setValues([[2]]);
    await sheet.range('C1:C2').setFormulas([['SUMIF(A1:A2,">0",B100)'], ['SUMIF(A1:B1,">0",Z1)']]);
    const initial = await sheet.range('C1:C2').getValues();
    await book.transaction({}, async (tx) => {
      const target = tx.getSheetByName('Projected bounds')!;
      await target.range('A2').setValues([[0]]);
      await target.range('B1').setValues([[0]]);
    });
    const edited = await sheet.range('C1:C2').getValues();
    await book.undo();
    const undone = await sheet.range('C1:C2').getValues();
    await book.redo();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheetByName('Projected bounds')!;
    return { initial, edited, undone, reopened: await sheet.range('C1:C2').getValues() };
  });
  expect(result).toEqual({
    initial: [[{ error: '#REF!' }], [{ error: '#REF!' }]],
    edited: [[2], [2]],
    undone: [[{ error: '#REF!' }], [{ error: '#REF!' }]],
    reopened: [[2], [2]],
  });
});

test('recalculates repeated and escaped wildcard criteria after edits and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1:A3').setValues([[''], ['*tail'], ['tail']]);
    await sheet
      .range('B1:B2')
      .setFormulas([['COUNTIF(A1:A3,"*****")'], ['COUNTIF(A1:A3,"**~***")']]);
    const initial = await sheet.range('B1:B2').getValues();
    await sheet.range('A3').setValues([['*next']]);
    const edited = await sheet.range('B1:B2').getValues();
    await app.open(book.id);
    book = app.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('B1:B2').getValues();
    await book.undo();
    const undone = await sheet.range('B1:B2').getValues();
    return { initial, edited, reopened, undone };
  });
  expect(result).toEqual({
    initial: [[3], [1]],
    edited: [[3], [2]],
    reopened: [[3], [2]],
    undone: [[3], [1]],
  });
});
