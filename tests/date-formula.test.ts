import { describe, expect, it } from 'vitest';
import { evaluateFormula, parseFormula } from '@opensheetjs/formula';
import { createWorkbookFile, Workbook, openWorkbook, FileReader } from '@opensheetjs/core';
import { fromSheetJS, toSheetJS } from '@opensheetjs/adapter-sheetjs';
import { date1904Corpus } from './fixtures/formula-corpus.js';
import { readFileSync } from 'node:fs';
import * as XLSX from 'xlsx';

describe('workbook date system', () => {
  it('recalculates an independently saved 1904 workbook after imported caches expire', async () => {
    const input = XLSX.read(
      readFileSync(
        new URL(
          '../benchmarks/formula-reader-2026-10-23-dates-1904.recalculated.xlsx',
          import.meta.url,
        ),
      ),
      { type: 'buffer', xlfn: true },
    );
    expect(input.Workbook?.WBProps?.date1904).toBe(true);
    const file = fromSheetJS(input).file;
    expect(file.dateSystem).toBe('1904');
    const book = await new Workbook(file).ready();
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('D1').setValues([['invalidate imported caches']]);
      for (const [expression, expected] of [
        ['DATE(2024,2,29)', 43889],
        ['YEAR(0)', 1904],
        ['DAY(0)', 1],
      ] as const) {
        const key = Object.keys(input.Sheets.Semantics).find(
          (key) => /^C\d+$/.test(key) && input.Sheets.Semantics[key].v === expression,
        )!;
        expect(key).toBeDefined();
        expect(await sheet.range('B' + key.slice(1)).getValues()).toEqual([[expected]]);
      }
    } finally {
      await book.close();
    }
  });
  it.each(date1904Corpus)('evaluates %s in 1904 mode', (expression, expected) => {
    expect(evaluateFormula(parseFormula(expression), () => null, undefined, '1904')).toEqual(
      expected,
    );
  });

  it('preserves 1904 semantics across Worker, export reader, edit/undo and reopen', async () => {
    const file = createWorkbookFile();
    file.dateSystem = '1904';
    let book = await new Workbook(file).ready();
    try {
      const sheet = book.getSheets()[0];
      await sheet
        .range(`A1:A${date1904Corpus.length}`)
        .setFormulas(date1904Corpus.map(([expression]) => [expression]));
      const expected = date1904Corpus.map(([, value]) => [value]);
      expect(await sheet.range(`A1:A${expected.length}`).getValues()).toEqual(expected);
      await sheet.range('B1').setValues([[2024]]);
      await sheet.range('C1').setFormulas([['DATE(B1,2,29)']]);
      await sheet.range('C1').setStyle({ numberFormat: 'yyyy-mm-dd' });
      expect(await sheet.range('C1').getDisplayValues()).toEqual([['2024-02-29']]);
      await sheet.range('B1').setValues([[2025]]);
      expect(await sheet.range('C1').getDisplayValues()).toEqual([['2025-03-01']]);
      await book.undo();
      const saved = await book.toJSON();
      const reader = new FileReader(saved);
      try {
        expect(reader.getSheetById(sheet.id)!.range(`A1:A${expected.length}`).getValues()).toEqual(
          expected,
        );
        expect(reader.getSheetById(sheet.id)!.range('C1').getDisplayValues()).toEqual([
          ['2024-02-29'],
        ]);
      } finally {
        reader.dispose();
      }
      const output = toSheetJS(saved);
      expect(output.workbook.Workbook?.WBProps?.date1904).toBe(true);
      expect(fromSheetJS(output.workbook).file.dateSystem).toBe('1904');
      const id = book.id;
      await book.close();
      book = await openWorkbook(id);
      expect(await book.getSheets()[0].range(`A1:A${expected.length}`).getValues()).toEqual(
        expected,
      );
      expect(await book.getSheets()[0].range('C1').getValues()).toEqual([[43889]]);
    } finally {
      await book.close();
    }
  });
});
