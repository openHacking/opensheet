import { expect, it } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { evaluateFormula, offsetFormula, parseFormula, printFormula } from '@opensheetjs/formula';

it('preserves overflow errors and lazy branches when formulas are rewritten', () => {
  expect(printFormula(parseFormula('1e309'))).toBe('#NUM!');
  const formula = offsetFormula('IF(FALSE,1e309,A1)', 1, 0);
  expect(formula).toContain('#NUM!');
  expect(formula).not.toContain('INFINITY');
  expect(evaluateFormula(parseFormula(formula), () => 9)).toBe(9);
  const caught = offsetFormula('IFERROR(1e309,A1)', 1, 0);
  expect(evaluateFormula(parseFormula(caught), () => 9)).toBe(9);
});

it('retains overflow handling through copy, structural edits, undo and reopen', async () => {
  let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 5, columns: 2 }] });
  try {
    let sheet = book.getSheets()[0];
    await sheet.range('A1:A2').setValues([[5], [9]]);
    await sheet.range('B1').setFormulas([['IFERROR(1e309,A1)']]);
    expect(await sheet.range('B1').getValues()).toEqual([[5]]);
    await sheet.range('B1').copyTo('B2');
    expect(await sheet.range('B2').getValues()).toEqual([[9]]);
    expect((await sheet.range('B2').getFormulas())[0][0]).toContain('#NUM!');
    await sheet.insertRows(0);
    expect(await sheet.range('B3').getValues()).toEqual([[9]]);
    await book.undo();
    expect(await sheet.range('B2').getValues()).toEqual([[9]]);
    await book.redo();
    const id = book.id;
    await book.close();
    book = await openWorkbook(id, { cacheBytes: 0 });
    sheet = book.getSheets()[0];
    expect(await sheet.range('B3').getValues()).toEqual([[9]]);
    expect((await sheet.range('B3').getFormulas())[0][0]).toContain('#NUM!');
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});
