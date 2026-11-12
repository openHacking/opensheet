import { describe, expect, it, vi } from 'vitest';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { Workbook } from '@opensheetjs/core';
import { fromSheetJS } from '@opensheetjs/adapter-sheetjs';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';

describe('lookup bounds and dependencies', () => {
  it('bounds repeated long text normalization for literal, wildcard and binary lookups', async () => {
    const text = 'x'.repeat(32_767);
    for (const expression of [
      'XLOOKUP("needle",A1:A31,B1:B31,"missing",0)',
      'MATCH("needle",A1:A31,0)',
      'XLOOKUP("n*",A1:A31,B1:B31,"missing",2)',
    ]) {
      expect(
        evaluateFormula(parseFormula(expression), () => text),
        expression,
      ).toEqual({ error: '#LIMIT!' });
    }
    expect(
      evaluateFormula(parseFormula('XLOOKUP("needle",A1:A30,B1:B30,"missing",0)'), () => text),
    ).toBe('missing');
    const target = 'y'.repeat(32_767);
    const binary = 'XLOOKUP(A1,B1:B65536,C1:C65536,"missing",0,2)';
    expect(
      evaluateFormula(parseFormula(binary), ({ column }) => (column === 0 ? target : text)),
    ).toEqual({ error: '#LIMIT!' });
    const expression = 'XLOOKUP("needle",A1:A31,B1:B31,"missing",0)';
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
    expect(await calculation.value('sheet', 0, 2)).toEqual({ error: '#LIMIT!' });
  });
  it('preserves wildcard lookup order, case handling, types and encountered errors', () => {
    const values: Array<string | number | { error: string }> = [7, 'aBc', 'AX', '*'];
    const read = ({ row, column }: { row: number; column: number }) =>
      column === 0 ? values[row] : row * 10;
    expect(evaluateFormula(parseFormula('MATCH("A*",A1:A4,0)'), read)).toBe(2);
    expect(evaluateFormula(parseFormula('VLOOKUP("A*",A1:B4,2,FALSE())'), read)).toBe(10);
    expect(evaluateFormula(parseFormula('XLOOKUP("a*",A1:A4,B1:B4,-1,2,-1)'), read)).toBe(20);
    expect(evaluateFormula(parseFormula('MATCH("~*",A1:A4,0)'), read)).toBe(4);
    values[0] = { error: '#DIV/0!' };
    expect(evaluateFormula(parseFormula('MATCH("A*",A1:A4,0)'), read)).toEqual({
      error: '#DIV/0!',
    });
  });
  it('imports an independently saved prefixed XLOOKUP and recalculates its edited result', async () => {
    const input = XLSX.read(
      readFileSync(
        new URL(
          '../benchmarks/formula-reader-2026-10-23-lookup.recalculated.xlsx',
          import.meta.url,
        ),
      ),
      { type: 'buffer', xlfn: true },
    );
    const sheet = input.Sheets.Semantics;
    const expression = 'XLOOKUP("alpha",D1:D5,E1:E5,"missing",0,-1)';
    const key = Object.keys(sheet).find(
      (key) => /^C\d+$/.test(key) && sheet[key].v === expression,
    )!;
    expect(key).toBeDefined();
    const address = 'B' + key.slice(1);
    expect(sheet[address].f).toContain('_xlfn.XLOOKUP');
    expect(sheet[address].v).toBe(30);
    const book = await new Workbook(fromSheetJS(input).file).ready();
    try {
      await book
        .getSheets()[0]
        .range('E3')
        .setValues([[90]]);
      expect(await book.getSheets()[0].range(address).getValues()).toEqual([[90]]);
    } finally {
      await book.close();
    }
  });
  it('reads only the selected INDEX cell and returns horizontal positions correctly', () => {
    const read = vi.fn(({ row, column }: { row: number; column: number }) => row * 10 + column);
    expect(evaluateFormula(parseFormula('INDEX(A1:C100,40,2)'), read)).toBe(391);
    expect(read).toHaveBeenCalledTimes(1);
    read.mockClear();
    expect(evaluateFormula(parseFormula('INDEX(A1:C1,2)'), read)).toBe(1);
    expect(evaluateFormula(parseFormula('MATCH(2,A1:C1,0)'), read)).toBe(3);
    expect(evaluateFormula(parseFormula('XLOOKUP(1,A1:C1,A2:C2)'), read)).toBe(11);
  });

  it('validates unsupported modes/shapes without traversing and bounds unsuccessful scans', () => {
    expect(evaluateFormula(parseFormula('MATCH(A1,B1:B2,0)'), () => 'a'.repeat(32768))).toEqual({
      error: '#LIMIT!',
    });
    const read = vi.fn(() => 1);
    for (const expression of [
      'MATCH(1,A1:A100,3)',
      'XLOOKUP(1,A1:A100,B1:B100,0,2,2)',
      'XLOOKUP(1,A1:A100,B1:B99)',
      'INDEX(A1:A100,101)',
    ]) {
      expect(evaluateFormula(parseFormula(expression), read)).toEqual({
        error: expression.startsWith('INDEX') ? '#REF!' : '#VALUE!',
      });
      expect(read).not.toHaveBeenCalled();
    }
    expect(evaluateFormula(parseFormula('MATCH(2,A1:A100000,0)'), read, 20)).toEqual({
      error: '#LIMIT!',
    });
    expect(read.mock.calls.length).toBeLessThan(20);
    read.mockClear();
    expect(evaluateFormula(parseFormula('MATCH(2,A1:A100000,1)'), read, 20)).toEqual({
      error: '#LIMIT!',
    });
    expect(read.mock.calls.length).toBeLessThan(20);
    expect(
      evaluateFormula(parseFormula('IFERROR(MATCH(2,A1:A2,0),"error")'), () => ({
        error: '#DIV/0!',
      })),
    ).toBe('error');
  });

  it('uses logarithmic reads for ascending and descending binary XLOOKUP', () => {
    for (const descending of [false, true]) {
      const read = vi.fn(({ row, column }: { row: number; column: number }) => {
        const value = descending ? 100000 - row : row + 1;
        return column === 0 ? value : value * 2;
      });
      const direction = descending ? -2 : 2;
      for (const [target, mode, expected] of [
        [50000, 0, 100000],
        [50000.5, -1, 100000],
        [50000.5, 1, 100002],
        [0, -1, -1],
        [100001, 1, -1],
      ]) {
        read.mockClear();
        expect(
          evaluateFormula(
            parseFormula(`XLOOKUP(${target},A1:A100000,B1:B100000,-1,${mode},${direction})`),
            read,
          ),
        ).toBe(expected);
        expect(read.mock.calls.length).toBeLessThan(22);
      }
    }
  });

  it('recalculates binary lookups after key edits, undo, redo and reopen', async () => {
    let book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B5').setValues([
        [1, 10],
        [2, 20],
        [3, 30],
        [4, 40],
        [5, 50],
      ]);
      await sheet
        .range('D1:D2')
        .setFormulas([
          ['XLOOKUP(2.5,A1:A5,B1:B5,"missing",-1,2)'],
          ['XLOOKUP(0,A1:A5,B1:B5,"missing",0,2)'],
        ]);
      expect(await sheet.range('D1:D2').getValues()).toEqual([[20], ['missing']]);
      await sheet.range('A3').setValues([[2.25]]);
      expect(await sheet.range('D1:D2').getValues()).toEqual([[30], ['missing']]);
      await book.undo();
      expect(await sheet.range('D1').getValues()).toEqual([[20]]);
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(await book.getSheets()[0].range('D1:D2').getValues()).toEqual([[30], ['missing']]);
    } finally {
      await book.close();
    }
  });

  it('supports bounded ordered approximate MATCH and VLOOKUP, rejecting unsorted and mixed keys', async () => {
    let book = await createWorkbook();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B5').setValues([
        [1, 10],
        [2, 20],
        [2, 21],
        [4, 40],
        [8, 80],
      ]);
      await sheet.range('D1:D5').setValues([[8], [4], [2], [2], [1]]);
      await sheet
        .range('F1:F8')
        .setFormulas([
          ['MATCH(3,A1:A5)'],
          ['MATCH(2,A1:A5,1)'],
          ['MATCH(3,D1:D5,-1)'],
          ['MATCH(0,A1:A5,1)'],
          ['VLOOKUP(3,A1:B5,2)'],
          ['VLOOKUP(2,A1:B5,2,TRUE)'],
          ['VLOOKUP(2,A1:B5,2,FALSE)'],
          ['MATCH(3,A1:A5,0)'],
        ]);
      const expected = [[3], [3], [2], [{ error: '#N/A' }], [21], [21], [20], [{ error: '#N/A' }]];
      expect(await sheet.range('F1:F8').getValues()).toEqual(expected);
      await sheet.range('A2').setValues([[9]]);
      expect(await sheet.range('F1').getValues()).toEqual([[{ error: '#VALUE!' }]]);
      await book.undo();
      await sheet.range('A5').setValues([[0]]);
      expect(await sheet.range('F1').getValues()).toEqual([[{ error: '#VALUE!' }]]);
      await book.undo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(await book.getSheets()[0].range('F1:F8').getValues()).toEqual(expected);
      await book
        .getSheets()[0]
        .range('A2')
        .setValues([['text']]);
      expect(await book.getSheets()[0].range('F1').getValues()).toEqual([[{ error: '#VALUE!' }]]);
    } finally {
      await book.close();
    }
  });

  it('invalidates selected and duplicate matches through edits, undo, insertion and reopen', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Totals' }] });
    try {
      const [data, totals] = book.getSheets();
      await data.range('A1:B3').setValues([
        ['key', 10],
        ['other', 20],
        ['KEY', 30],
      ]);
      await totals
        .range('A1:A2')
        .setFormulas([
          ['XLOOKUP("key",Data!A1:A3,Data!B1:B3,"missing",0,-1)'],
          ['VLOOKUP("key",Data!A1:B3,2,FALSE)'],
        ]);
      expect(await totals.range('A1:A2').getValues()).toEqual([[30], [10]]);
      await data.range('A3').setValues([['different']]);
      expect(await totals.range('A1:A2').getValues()).toEqual([[10], [10]]);
      await book.undo();
      await data.insertRows(0);
      expect(await totals.range('A1:A2').getValues()).toEqual([[30], [10]]);
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(await book.getSheets()[1].range('A1:A2').getValues()).toEqual([[30], [10]]);
      await book
        .getSheets()[0]
        .range('B4')
        .setValues([[50]]);
      expect(await book.getSheets()[1].range('A1:A2').getValues()).toEqual([[50], [10]]);
    } finally {
      await book.close();
    }
  });
});
