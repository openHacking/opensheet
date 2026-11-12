import { expect, test } from '@playwright/test';

test('cycles propagate through IFERROR and recover after edits, history and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const first = await book.addSheet('Cycle A', { rows: 10, columns: 4 });
    const second = await book.addSheet('Cycle B', { rows: 10, columns: 2 });
    await first
      .range('A1:D1')
      .setFormulas([
        [
          "IFERROR('Cycle B'!A1,99)",
          'IFERROR(1/0,42)',
          'IFERROR(CONCAT(1/0,C1),99)',
          'IFERROR(TEXTJOIN(1/0,TRUE,D1),99)',
        ],
      ]);
    await second.range('A1:B1').setFormulas([["'Cycle A'!A1+1", 'IF(FALSE,A1,3)']]);
    const read = async () => ({
      a: (await first.range('A1').getValues())[0][0],
      b: (await second.range('A1').getValues())[0][0],
      ordinary: (await first.range('B1').getValues())[0][0],
      lazy: (await second.range('B1').getValues())[0][0],
      concat: (await first.range('C1').getValues())[0][0],
      join: (await first.range('D1').getValues())[0][0],
    });
    const cycle = await read();
    await second.range('A1').setValues([[7]]);
    const repaired = await read();
    await book.undo();
    const undone = await read();
    await book.redo();
    const redone = await read();
    await app.open(book.id);
    const reopened = await app.getWorkbook().getSheetByName('Cycle A')!.range('A1').getValues();
    return { cycle, repaired, undone, redone, reopened };
  });
  const cycle = {
    a: { error: '#CYCLE!' },
    b: { error: '#CYCLE!' },
    ordinary: 42,
    lazy: 3,
    concat: { error: '#CYCLE!' },
    join: { error: '#CYCLE!' },
  };
  const repaired = { ...cycle, a: 7, b: 7 };
  expect(result).toEqual({ cycle, repaired, undone: cycle, redone: repaired, reopened: [[7]] });
});

test('mixed-type comparisons and boolean concatenation recalculate after edits and reopen', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = await book.addSheet('Comparison', { rows: 4, columns: 6 });
    await sheet.range('A1:B1').setValues([['x', true]]);
    await sheet.range('C1:F1').setFormulas([['A1<B1', 'B1>A1', 'B1<A1', 'B1&"/"&A1']]);
    const before = await sheet.range('C1:F1').getValues();
    await sheet.range('B1').setValues([[false]]);
    const changed = await sheet.range('C1:F1').getValues();
    await book.undo();
    const undone = await sheet.range('C1:F1').getValues();
    await book.redo();
    await app.open(book.id);
    const reopened = await app
      .getWorkbook()
      .getSheetByName('Comparison')!
      .range('C1:F1')
      .getValues();
    return { before, changed, undone, reopened };
  });
  expect(result).toEqual({
    before: [[true, true, false, 'TRUE/x']],
    changed: [[true, true, false, 'FALSE/x']],
    undone: [[true, true, false, 'TRUE/x']],
    reopened: [[true, true, false, 'FALSE/x']],
  });
});

test('early errors skip later dependencies but editing the first operand recalculates them', async ({
  page,
}) => {
  await page.goto('/#budget');
  await page.waitForFunction(() => !!window.opensheet?.selection.get());
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = await book.addSheet('Short circuit', { rows: 10, columns: 4 });
    await sheet.range('A1:D1').setFormulas([['1/0', 'B1', 'A1+B1', 'ROUND(A1,B1)']]);
    const read = () => sheet.range('C1:D1').getValues();
    const before = await read();
    await sheet.range('A1').setValues([[1]]);
    const after = await read();
    await sheet.range('B1').setValues([[2]]);
    const repaired = await read();
    return { before, after, repaired };
  });
  expect(result).toEqual({
    before: [[{ error: '#DIV/0!' }, { error: '#DIV/0!' }]],
    after: [[{ error: '#CYCLE!' }, { error: '#CYCLE!' }]],
    repaired: [[3, 1]],
  });
});
