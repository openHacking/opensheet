import { describe, expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook, type Workbook } from '@opensheetjs/core';
import {
  evaluateFormula,
  parseFormula,
  functionDefinitions,
  supportedFunctions,
} from '@opensheetjs/formula';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';
import { FormulaCalculation } from '../packages/core/src/calculation.js';
import { createSheet } from '../packages/core/src/model.js';
import type { SheetSnapshot } from '../packages/core/src/types.js';
import {
  formulaCorpus,
  formulaInputs,
  formulaLookupInputs,
  orderedLookupInputs,
} from './fixtures/formula-corpus.js';

describe('independent formula contract corpus', () => {
  it.each([Infinity, -Infinity, NaN])(
    'treats nonfinite external reference %s as an error at every use',
    (value) => {
      const read = vi.fn(() => value);
      expect(evaluateFormula(parseFormula('A1'), read)).toEqual({ error: '#NUM!' });
      expect(evaluateFormula(parseFormula('MIN(A1:A2)'), read)).toEqual({ error: '#NUM!' });
      expect(evaluateFormula(parseFormula('IFERROR(A1,7)'), read)).toBe(7);
      read.mockClear();
      expect(evaluateFormula(parseFormula('IF(FALSE,A1,3)'), read)).toBe(3);
      expect(read).not.toHaveBeenCalled();
    },
  );
  it('does not hide a later calculation limit behind AVERAGE invalid text', async () => {
    const expression = 'IFERROR(AVERAGE("bad",A1),7)';
    expect(evaluateFormula(parseFormula(expression), () => ({ error: '#LIMIT!' }))).toEqual({
      error: '#LIMIT!',
    });
    const calculation = new AsyncCalculation(
      async (_sheet, _row, column) => ({
        rowId: 0,
        columnId: column,
        input:
          column === 0
            ? { type: 'error' as const, code: '#LIMIT!' }
            : { type: 'formula' as const, expression },
      }),
      () => undefined,
    );
    expect(await calculation.value('sheet', 0, 1)).toEqual({ error: '#LIMIT!' });
  });
  it.each([
    ['MIN("  +1.25e2  ",200)', 125],
    ['MAX("-3",-7)', -3],
    ['MIN("0x10",1)', { error: '#VALUE!' }],
    ['MAX("Infinity",1)', { error: '#VALUE!' }],
    ['MIN("   ",1)', { error: '#VALUE!' }],
    ['COUNT("7",TRUE,"bad")', 2],
  ] as const)('applies bounded direct numeric coercion in %s', (expression, expected) => {
    expect(evaluateFormula(parseFormula(expression), () => null)).toEqual(expected);
  });
  it.each(formulaCorpus)('%s [%s]', (expression, expected) => {
    const value = evaluateFormula(parseFormula(expression), (ref) =>
      ref.column === 0
        ? (formulaInputs[ref.row] ?? null)
        : ref.column >= 6
          ? (orderedLookupInputs[ref.row]?.[ref.column - 6] ?? null)
          : (formulaLookupInputs[ref.row]?.[ref.column - 3] ?? null),
    );
    expect(value).toEqual(expected);
  });

  it.each([undefined, 0])(
    'matches the corpus in the worker and after reopen (cache %s)',
    async (cacheBytes) => {
      let book: Workbook = await createWorkbook(
        { sheets: [{ name: 'Semantics', rows: Math.max(100, formulaCorpus.length) }] },
        { cacheBytes },
      );
      try {
        const sheet = book.getSheets()[0];
        await sheet.range('A1:A7').setValues([[true], ['7'], [2], [null], [''], [null], [false]]);
        await sheet.range('D1:E5').setValues(formulaLookupInputs);
        await sheet.range('G1:I5').setValues(orderedLookupInputs);
        await sheet.range('A6').setFormulas([['1/0']]);
        await sheet
          .range(`B1:B${formulaCorpus.length}`)
          .setFormulas(formulaCorpus.map(([expression]) => [expression]));
        const expected = formulaCorpus.map(([, value]) => [value]);
        expect(await sheet.range(`B1:B${formulaCorpus.length}`).getValues()).toEqual(expected);
        const id = book.id;
        await book.close();
        book = await openWorkbook(id, { cacheBytes });
        expect(await book.getSheets()[0].range(`B1:B${formulaCorpus.length}`).getValues()).toEqual(
          expected,
        );
      } finally {
        await book.close();
      }
    },
    30000,
  );

  it('rejects unknown functions before visiting arguments in both evaluators', async () => {
    const read = vi.fn(() => null);
    expect(evaluateFormula(parseFormula('UNSUPPORTED(A1:A99999)'), read)).toEqual({
      error: '#NAME?',
    });
    expect(read).not.toHaveBeenCalled();
    const asyncRead = vi.fn(async () => ({
      rowId: 0,
      columnId: 0,
      input: { type: 'formula' as const, expression: 'UNSUPPORTED(B1:B99999)' },
    }));
    const calculation = new AsyncCalculation(asyncRead, () => undefined);
    expect(await calculation.value('sheet', 0, 0)).toEqual({ error: '#NAME?' });
    expect(asyncRead).toHaveBeenCalledTimes(1);
  });

  it.each(['A1+B1', 'ROUND(A1,B1)', 'LEFT(A1,B1)'])(
    'stops evaluating %s after the first error in pure and asynchronous calculation',
    async (expression) => {
      const pureRead = vi.fn((ref: { column: number }) => {
        if (ref.column === 1) throw new Error('unreachable');
        return { error: '#DIV/0!' };
      });
      expect(evaluateFormula(parseFormula(expression), pureRead)).toEqual({ error: '#DIV/0!' });
      expect(pureRead).toHaveBeenCalledTimes(1);
      const read = vi.fn(async (_sheet: string, _row: number, column: number) => {
        if (column === 1) throw new Error('unreachable');
        return {
          rowId: 0,
          columnId: column,
          input:
            column === 2
              ? { type: 'formula' as const, expression }
              : { type: 'error' as const, code: '#DIV/0!' },
        };
      });
      const calculation = new AsyncCalculation(read, () => undefined);
      expect(await calculation.value('sheet', 0, 2)).toEqual({ error: '#DIV/0!' });
      expect(read.mock.calls.map((call) => call[2])).toEqual([2, 0]);
    },
  );

  it('does not turn exhausted work or circular dependencies into successful counts', async () => {
    expect(evaluateFormula(parseFormula('COUNT(SUM(A1:A100))'), () => 1, 8)).toEqual({
      error: '#LIMIT!',
    });
    for (const error of ['#LIMIT!', '#CYCLE!']) {
      expect(evaluateFormula(parseFormula('COUNT(A1)'), () => ({ error }))).toEqual({ error });
      const calculation = new AsyncCalculation(
        async (_sheet, _row, column) => ({
          rowId: 0,
          columnId: column,
          input:
            column === 0
              ? { type: 'formula', expression: 'COUNT(B1)' }
              : { type: 'error', code: error },
        }),
        () => undefined,
      );
      expect(await calculation.value('sheet', 0, 0)).toEqual({ error });
    }
  });

  it('does not cache depth exhaustion or let IFERROR conceal it', async () => {
    expect(evaluateFormula(parseFormula('IFERROR(A1,99)'), () => ({ error: '#LIMIT!' }))).toEqual({
      error: '#LIMIT!',
    });
    const calculation = new AsyncCalculation(
      async (_sheet, row) => ({
        rowId: row,
        columnId: 0,
        input:
          row === 0
            ? { type: 'number', value: 1 }
            : { type: 'formula', expression: row === 129 ? 'IFERROR(A129,99)' : `A${row}+1` },
      }),
      () => undefined,
    );
    expect(await calculation.value('sheet', 129, 0)).toEqual({ error: '#LIMIT!' });
    expect(await calculation.value('sheet', 128, 0)).toBe(129);
    expect(await calculation.value('sheet', 129, 0)).toBe(129);
    const cells = Object.fromEntries(
      Array.from({ length: 130 }, (_, row) => [
        `${row}:0`,
        {
          rowId: row,
          columnId: 0,
          input:
            row === 0
              ? { type: 'number', value: 1 }
              : { type: 'formula', expression: row === 129 ? 'IFERROR(A129,99)' : `A${row}+1` },
        },
      ]),
    );
    const sheet = {
      rowOrder: Array.from({ length: 130 }, (_, row) => row),
      columnOrder: [0],
      cells,
    } as SheetSnapshot;
    const synchronous = new FormulaCalculation(
      () => sheet,
      () => undefined,
    );
    expect(synchronous.value('sheet', 129, 0)).toEqual({ error: '#LIMIT!' });
    expect(synchronous.value('sheet', 128, 0)).toBe(129);
    expect(synchronous.value('sheet', 129, 0)).toBe(129);
  });

  it('propagates circular dependency failures through IFERROR independently of read order', async () => {
    expect(evaluateFormula(parseFormula('IFERROR(A1,99)'), () => ({ error: '#CYCLE!' }))).toEqual({
      error: '#CYCLE!',
    });
    for (const first of [0, 1]) {
      let cyclic = true;
      const inputs = (column: number) =>
        column === 0
          ? { type: 'formula' as const, expression: 'IFERROR(B1,99)' }
          : cyclic
            ? { type: 'formula' as const, expression: 'A1+1' }
            : { type: 'number' as const, value: 7 };
      const asyncCalc = new AsyncCalculation(
        async (_sheet, _row, column) => ({
          rowId: 0,
          columnId: column,
          input: inputs(column),
        }),
        () => undefined,
      );
      const sheet = createSheet('Cycle', 1, 2);
      sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: inputs(0) };
      sheet.cells['0:1'] = { rowId: 0, columnId: 1, input: inputs(1) };
      const syncCalc = new FormulaCalculation(
        () => sheet,
        () => undefined,
      );
      for (const column of [first, 1 - first]) {
        expect(await asyncCalc.value('sheet', 0, column)).toEqual({ error: '#CYCLE!' });
        expect(syncCalc.value('sheet', 0, column)).toEqual({ error: '#CYCLE!' });
      }
      cyclic = false;
      sheet.cells['0:1'].input = inputs(1);
      asyncCalc.invalidate([
        { sheetId: 'sheet', startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 },
      ]);
      syncCalc.clear();
      expect(await asyncCalc.value('sheet', 0, 0)).toBe(7);
      expect(syncCalc.value('sheet', 0, 0)).toBe(7);
    }
    expect(evaluateFormula(parseFormula('IF(FALSE,A1,42)'), () => ({ error: '#CYCLE!' }))).toBe(42);
  });

  it.each(['#LIMIT!', '#CYCLE!'])(
    'does not mask %s behind an earlier text aggregation error',
    async (internalError) => {
      for (const expression of [
        'IFERROR(CONCAT(A1,B1),99)',
        'IFERROR(TEXTJOIN(A1,TRUE,B1),99)',
        'IFERROR(TEXTJOIN(1/0,TRUE,B1),99)',
        'IFERROR(TEXTJOIN(A1,B1,"x"),99)',
      ]) {
        const read = (column: number) => ({ error: column === 0 ? '#DIV/0!' : internalError });
        expect(evaluateFormula(parseFormula(expression), (ref) => read(ref.column))).toEqual({
          error: internalError,
        });
        const calculation = new AsyncCalculation(
          async (_sheet, _row, column) => ({
            rowId: 0,
            columnId: column,
            input:
              column === 2
                ? { type: 'formula', expression }
                : { type: 'error', code: read(column).error },
          }),
          () => undefined,
        );
        expect(await calculation.value('sheet', 0, 2)).toEqual({ error: internalError });
      }
    },
  );

  it('covers every registered function and protects its shared help/arity metadata', () => {
    expect(supportedFunctions).toEqual(functionDefinitions.map(({ name }) => name));
    for (const definition of functionDefinitions) {
      expect(Object.isFrozen(definition)).toBe(true);
      expect(
        formulaCorpus.some(([expression]) => expression.startsWith(`${definition.name}(`)),
      ).toBe(true);
      expect(definition.signature.startsWith(`${definition.name}(`)).toBe(true);
      expect(
        evaluateFormula(
          {
            kind: 'call',
            name: definition.name,
            args: definition.minArgs ? [] : [{ kind: 'literal', value: 1 }],
          },
          () => null,
        ),
      ).toEqual({ error: '#VALUE!' });
    }
    expect(Object.isFrozen(functionDefinitions)).toBe(true);
    expect(Object.isFrozen(supportedFunctions)).toBe(true);
  });
});
