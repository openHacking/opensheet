import { describe, expect, it, vi } from 'vitest';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { criterion } from '../packages/formula/src/criteria.js';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';

describe('conditional aggregation', () => {
  it('matches escaped wildcards, Unicode and multiline text without regex interpretation', () => {
    for (const [pattern, value, matches] of [
      ['a*', 'ALPHA', true],
      ['a?c', 'a🙂c', true],
      ['a?c', 'ac', false],
      ['~*', '*', true],
      ['~?', '?', true],
      ['~~', '~', true],
      ['~*', 'anything', false],
      ['a*b', 'a\nline\nb', true],
      ['[a]', 'a', false],
      ['[a]', '[a]', true],
      ['<>a*', 'beta', true],
      ['a*b*c', 'abXYZbYc', true],
      ['**~***', '*suffix', true],
      ['**~***', 'suffix', false],
    ] as const)
      expect(criterion(pattern)(value), pattern).toBe(matches);
  });

  it('collapses consecutive stars without changing escaped-star semantics', () => {
    const matches = criterion('*'.repeat(32_767));
    for (let i = 0; i < 10_000; i++) expect(matches('')).toBe(true);
    expect(matches('value')).toBe(true);
    const read = () => '';
    expect(evaluateFormula(parseFormula('COUNTIF(A1:A100,"*****")'), read)).toBe(100);
  });

  it('bounds literal and numeric-text comparisons as well as wildcard matching', async () => {
    const text = 'x'.repeat(32_767);
    const expression = 'COUNTIF(A1:A31,"needle")';
    const reads = vi.fn(() => text);
    expect(evaluateFormula(parseFormula(expression), reads)).toEqual({ error: '#LIMIT!' });
    expect(reads).toHaveBeenCalledTimes(31);
    expect(evaluateFormula(parseFormula('COUNTIF(A1:A30,"needle")'), () => text)).toBe(0);
    const numericText = '0'.repeat(32_767);
    expect(evaluateFormula(parseFormula('COUNTIF(A1:A31,1)'), () => numericText)).toEqual({
      error: '#LIMIT!',
    });
    expect(evaluateFormula(parseFormula('COUNTIFS(A1:A16,">a",A1:A16,">a")'), () => text)).toEqual({
      error: '#LIMIT!',
    });
    expect(evaluateFormula(parseFormula('COUNTIFS(A1:A16,">z",A1:A16,">a")'), () => text)).toBe(0);
    const calculation = new AsyncCalculation(
      async (_sheet, row, column) => ({
        rowId: row,
        columnId: column,
        input:
          row === 0 && column === 1
            ? { type: 'formula', expression }
            : { type: 'string', value: text },
      }),
      () => undefined,
    );
    expect(await calculation.value('sheet', 0, 1)).toEqual({ error: '#LIMIT!' });
  });

  it('shares text work across nested conditional and lookup calls in one formula', async () => {
    const text = 'x'.repeat(32767);
    const inner = 'COUNTIF(A1:A20,"needle")';
    const expressions = [
      `SUM(${inner},${inner})`,
      `IFERROR(SUM(${inner},${inner}),123)`,
      `SUM(${inner},XLOOKUP("needle",A1:A20,B1:B20,0,0))`,
    ];
    for (const expression of expressions) {
      expect(
        evaluateFormula(parseFormula(expression), () => text),
        expression,
      ).toEqual({ error: '#LIMIT!' });
      const calculation = new AsyncCalculation(
        async (_sheet, row, column) => ({
          rowId: row,
          columnId: column,
          input:
            row === 0 && column === 2
              ? { type: 'formula', expression }
              : { type: 'string', value: text },
        }),
        () => undefined,
      );
      expect(await calculation.value('sheet', 0, 2), expression).toEqual({ error: '#LIMIT!' });
    }
    expect(evaluateFormula(parseFormula(`IF(FALSE(),SUM(${inner},${inner}),7)`), () => text)).toBe(
      7,
    );
  });

  it('rejects mismatched shapes before reading and bounds cell and wildcard work', () => {
    const work = { remaining: 11 };
    const matches = criterion('a*', work);
    expect(matches('aaa')).toBe(true);
    expect(() => matches('aaa')).toThrow('#LIMIT!');
    const read = vi.fn((_ref: { row: number; column: number }) => 1);
    expect(evaluateFormula(parseFormula('SUMIFS(A1:A3,B1:B2,1)'), read)).toEqual({
      error: '#VALUE!',
    });
    expect(read).not.toHaveBeenCalled();
    expect(evaluateFormula(parseFormula('COUNTIF(A1:A100000,1)'), read, 20)).toEqual({
      error: '#LIMIT!',
    });
    expect(read.mock.calls.length).toBeLessThan(20);
    expect(() => criterion('*' + 'a'.repeat(2000) + 'b')('a'.repeat(2000) + 'c')).toThrow(
      '#LIMIT!',
    );
    for (const error of ['#CYCLE!', '#LIMIT!'])
      expect(evaluateFormula(parseFormula('COUNTIF(A1:A2,1)'), () => ({ error }))).toEqual({
        error,
      });
  });

  it('does not read projected SUMIF and AVERAGEIF cells beyond A1 address limits', () => {
    const read = vi.fn((_ref: { row: number; column: number }) => 1);
    for (const expression of [
      'SUMIF(A1:A2,">0",A1048576)',
      'AVERAGEIF(A1:A2,">0",A1048576)',
      'SUMIF(A1:B1,">0",XFD1)',
      'AVERAGEIF(A1:B1,">0",XFD1)',
    ]) {
      read.mockClear();
      expect(evaluateFormula(parseFormula(expression), read)).toEqual({ error: '#REF!' });
      expect(read.mock.calls.every(([ref]) => ref.row < 1048576 && ref.column < 16384)).toBe(true);
    }
    expect(evaluateFormula(parseFormula('SUMIF(A1:A2,">0",A1048575)'), read)).toBe(2);
    expect(evaluateFormula(parseFormula('SUMIF(A1:A2,">1",A1048576)'), read)).toBe(0);
  });

  it('matches asynchronous calculation on a projected address overflow', async () => {
    const expression = 'SUMIF(A1:A2,">0",A1048576)';
    const read = vi.fn(async (_sheet: string, row: number, column: number) => ({
      rowId: row,
      columnId: column,
      input:
        row === 0 && column === 2
          ? ({ type: 'formula', expression } as const)
          : ({ type: 'number', value: 1 } as const),
    }));
    const calculation = new AsyncCalculation(read, () => undefined);
    expect(await calculation.value('data', 0, 2)).toEqual({ error: '#REF!' });
    expect(read.mock.calls.every(([, row, column]) => row < 1048576 && column < 16384)).toBe(true);
  });

  it('tracks projected ranges through edits, history, structural changes and reopen', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Totals' }] });
    try {
      const [data, totals] = book.getSheets();
      await data.range('A1:B3').setValues([
        ['yes', 10],
        ['no', 20],
        ['YES', 30],
      ]);
      await totals
        .range('A1:A3')
        .setFormulas([
          ['SUMIF(Data!A1:A3,"yes",Data!B1)'],
          ['AVERAGEIF(Data!A1:A3,"yes",Data!B1)'],
          ['COUNTIFS(Data!A1:A3,"yes",Data!B1:B3,">15")'],
        ]);
      expect(await totals.range('A1:A3').getValues()).toEqual([[40], [20], [1]]);
      await data.range('B3').setValues([[50]]);
      expect(await totals.range('A1:A3').getValues()).toEqual([[60], [30], [1]]);
      await book.undo();
      expect(await totals.range('A1').getValues()).toEqual([[40]]);
      await book.redo();
      await data.insertRows(0);
      expect(await totals.range('A1:A3').getValues()).toEqual([[60], [30], [1]]);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      expect(await book.getSheets()[1].range('A1:A3').getValues()).toEqual([[60], [30], [1]]);
      await book
        .getSheets()[0]
        .range('B4')
        .setValues([[70]]);
      expect(await book.getSheets()[1].range('A1:A3').getValues()).toEqual([[80], [40], [1]]);
    } finally {
      await book.close();
    }
  });
});
