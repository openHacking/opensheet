import { describe, expect, it } from 'vitest';
import fc from 'fast-check';
import {
  conditionalStyle,
  conditionalViewport,
  conditionalViewportWithFormulas,
  rewriteConditionalSheetReferences,
  transformConditionalRules,
  validateConditionalRules,
} from '../packages/core/src/conditional-format.js';
import { AsyncCalculation } from '../packages/core/src/async-calculation.js';

const range = { startRow: 1, endRow: 5, startColumn: 2, endColumn: 6 };
const numberRule = (id: string, first = 5) => ({
  id,
  range,
  condition: { kind: 'number', operator: 'gte', first },
  style: { color: '#ff0000', bold: true },
  stopIfTrue: false,
});

describe('conditional formatting model', () => {
  it('rewrites only explicit target sheet references after rename and removal', () => {
    const source = [
      {
        ...numberRule('formula'),
        condition: { kind: 'formula', expression: 'AND(lower!$A$1>A1,Other!B2="lower!A1")' },
      },
      numberRule('number'),
    ];
    const renamed = rewriteConditionalSheetReferences(source, 10, 10, 'Lower', "New 'name'");
    expect(renamed[0].condition).toEqual({
      kind: 'formula',
      expression: "AND(('New ''name'''!$A$1>A1),('Other'!B2=\"lower!A1\"))",
    });
    const removed = rewriteConditionalSheetReferences(renamed, 10, 10, "new 'NAME'", null);
    expect(removed[0].condition).toEqual({
      kind: 'formula',
      expression: 'AND((#REF!>A1),(\'Other\'!B2="lower!A1"))',
    });
    expect(renamed[1]).toEqual(source[1]);
    renamed[1].range.startRow = 0;
    expect(source[1].range.startRow).toBe(1);
    expect(source[0].condition).toEqual({
      kind: 'formula',
      expression: 'AND(lower!$A$1>A1,Other!B2="lower!A1")',
    });
    expect(() => rewriteConditionalSheetReferences([], 10, 10, '', null)).toThrow();
    expect(() =>
      rewriteConditionalSheetReferences(source, 10, 10, 'lower', 'x'.repeat(4096)),
    ).toThrow();
    expect(source[0].condition).toEqual({
      kind: 'formula',
      expression: 'AND(lower!$A$1>A1,Other!B2="lower!A1")',
    });
  });
  it('bounds viewport work before evaluating or returning results', () => {
    const rules = validateConditionalRules([numberRule('first')], 10, 10);
    const cells = [
      { row: 1, column: 2, value: 5, base: { color: '#111111' } },
      { row: 0, column: 0, value: 5, base: { color: '#111111' } },
    ];
    expect(conditionalViewport(rules, cells)).toEqual([
      { color: '#ff0000', bold: true },
      { color: '#111111' },
    ]);
    expect(cells[0].base).toEqual({ color: '#111111' });
    const manyRules = validateConditionalRules(
      Array.from({ length: 101 }, (_, i) => numberRule(`r${i}`)),
      10,
      10,
    );
    expect(() =>
      conditionalViewport(
        manyRules,
        Array.from({ length: 10000 }, () => cells[0]),
      ),
    ).toThrow('Conditional formatting viewport exceeds work budget');
    expect(() =>
      conditionalViewport(
        [],
        Array.from({ length: 10001 }, () => cells[0]),
      ),
    ).toThrow('Conditional formatting viewport exceeds work budget');
    const textRules = validateConditionalRules(
      Array.from({ length: 101 }, (_, i) => ({
        ...numberRule(`text${i}`),
        condition: {
          kind: 'text',
          operator: 'contains',
          value: 'a'.repeat(4096),
          caseSensitive: false,
        },
      })),
      10,
      10,
    );
    expect(() =>
      conditionalViewport(
        textRules,
        Array.from({ length: 100 }, () => ({ ...cells[0], value: 'a' })),
      ),
    ).toThrow('Conditional formatting viewport exceeds text work budget');
    expect(() => conditionalStyle([], 1, 2, 'a'.repeat(32769), {})).toThrow(
      'Conditional formatting value exceeds text budget',
    );
  });

  it('returns validated values rather than rereading mutable input accessors', () => {
    let reads = 0;
    const input = {
      ...numberRule('first'),
      get stopIfTrue() {
        return reads++ === 0 ? false : 'invalid';
      },
    };
    const result = validateConditionalRules([input], 10, 10);
    expect(result[0].stopIfTrue).toBe(false);
    expect(reads).toBe(1);
  });

  it('uses priority per property and stops only after a matching stop rule', () => {
    const input = [
      numberRule('first'),
      { ...numberRule('second'), style: { color: '#0000ff', background: '#00ff00', bold: false } },
      { ...numberRule('stop'), style: { italic: true }, stopIfTrue: true },
      { ...numberRule('last'), style: { underline: true } },
    ];
    const rules = validateConditionalRules(input, 10, 10);
    const base = { color: '#111111', background: '#ffffff', underline: false };
    expect(conditionalStyle(rules, 1, 2, 5, base)).toEqual({
      color: '#ff0000',
      background: '#00ff00',
      bold: true,
      italic: true,
      underline: false,
    });
    expect(conditionalStyle(rules, 1, 2, 4, base)).toEqual(base);
    expect(conditionalStyle(rules, 5, 2, 5, base)).toEqual(base);
    expect(base).toEqual({ color: '#111111', background: '#ffffff', underline: false });
    input[0].style.color = '#abcdef';
    expect(rules[0].style.color).toBe('#ff0000');
  });

  it('compares typed numeric and text values without coercing blanks or errors', () => {
    const rules = validateConditionalRules(
      [
        {
          ...numberRule('number'),
          condition: { kind: 'number', operator: 'between', first: 2, second: 4 },
          style: { bold: true },
        },
        {
          ...numberRule('text'),
          condition: { kind: 'text', operator: 'contains', value: 'ÉTÉ', caseSensitive: false },
          style: { italic: true },
        },
      ],
      10,
      10,
    );
    expect(conditionalStyle(rules, 1, 2, 2, {})).toEqual({ bold: true });
    expect(conditionalStyle(rules, 1, 2, 4, {})).toEqual({ bold: true });
    expect(conditionalStyle(rules, 1, 2, '2', {})).toEqual({});
    expect(conditionalStyle(rules, 1, 2, null, {})).toEqual({});
    expect(conditionalStyle(rules, 1, 2, { error: '#VALUE!' }, {})).toEqual({});
    expect(conditionalStyle(rules, 1, 2, 'un été chaud', {})).toEqual({ italic: true });
  });

  it.each([
    ['eq', 3, undefined, 3, true],
    ['neq', 3, undefined, 3, false],
    ['lt', 3, undefined, 2, true],
    ['lte', 3, undefined, 3, true],
    ['gt', 3, undefined, 4, true],
    ['gte', 3, undefined, 3, true],
    ['between', 2, 4, 3, true],
    ['notBetween', 2, 4, 4, false],
    ['notBetween', 2, 4, 5, true],
  ])('evaluates numeric operator %s', (operator, first, second, value, expected) => {
    const condition = {
      kind: 'number',
      operator,
      first,
      ...(second === undefined ? {} : { second }),
    };
    const rules = validateConditionalRules([{ ...numberRule('x'), condition }], 10, 10);
    expect(conditionalStyle(rules, 1, 2, value, {}).bold === true).toBe(expected);
  });

  it.each([
    ['equals', 'Alpha', 'alpha', true, false],
    ['notEquals', 'Alpha', 'alpha', false, false],
    ['contains', 'AlphaBeta', 'BETA', false, true],
    ['notContains', 'AlphaBeta', 'BETA', false, false],
    ['startsWith', 'AlphaBeta', 'alpha', false, true],
    ['endsWith', 'AlphaBeta', 'BETA', false, true],
    ['contains', '', '', true, true],
  ])('evaluates text operator %s', (operator, value, operand, caseSensitive, expected) => {
    const condition = { kind: 'text', operator, value: operand, caseSensitive };
    const rules = validateConditionalRules([{ ...numberRule('x'), condition }], 10, 10);
    expect(conditionalStyle(rules, 1, 2, value, {}).bold === true).toBe(expected);
  });

  it('rejects malformed rules, unbounded inputs and unsupported features', () => {
    for (const invalid of [
      {
        ...numberRule('x'),
        condition: { kind: 'number', operator: 'between', first: 5, second: 4 },
      },
      { ...numberRule('x'), condition: { kind: 'formula', expression: 'A1>' } },
      { ...numberRule('x'), style: { numberFormat: '0.00' } },
      { ...numberRule('x'), style: { bold: undefined } },
      { ...numberRule('x'), range: { ...range, endRow: 11 } },
      { ...numberRule('__proto__') },
    ])
      expect(() => validateConditionalRules([invalid], 10, 10)).toThrow();
    expect(() =>
      validateConditionalRules([numberRule('same'), numberRule('same')], 10, 10),
    ).toThrow();
    expect(() =>
      validateConditionalRules(
        Array.from({ length: 1025 }, (_, i) => numberRule(`x${i}`)),
        10,
        10,
      ),
    ).toThrow();
  });

  it('evaluates formula rules against a caller-owned staged calculation session', async () => {
    const calc = new AsyncCalculation(
      async (sheet, row, column) =>
        column === 0 && row < 2
          ? {
              rowId: row,
              columnId: column,
              input: { type: 'number', value: sheet === 'limits' ? 10 : row ? 11 : 5 },
            }
          : undefined,
      (name) => (name === 'Limits' ? 'limits' : undefined),
    );
    const rules = validateConditionalRules(
      [
        {
          id: 'formula',
          range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
          condition: { kind: 'formula', expression: 'AND(A1>Limits!$A$1,$A$1=5)' },
          style: { bold: true },
          stopIfTrue: false,
        },
      ],
      10,
      10,
    );
    const cells = [
      { row: 0, column: 0, value: 5, base: {} },
      { row: 1, column: 0, value: 11, base: {} },
    ];
    const session = calc.validationSession().forSheet('data');
    expect(await conditionalViewportWithFormulas(rules, cells, session.evaluateFormula)).toEqual([
      {},
      { bold: true },
    ]);
    expect(() => conditionalViewport(rules, cells)).toThrow('formula reader required');
    expect(await calc.value('data', 1, 0)).toBe(11);
  });

  it('handles formula errors and cancellation without returning partial appearance', async () => {
    const rules = validateConditionalRules(
      [{ ...numberRule('f'), condition: { kind: 'formula', expression: 'A1>0' } }],
      10,
      10,
    );
    const cells = [{ row: 1, column: 2, value: 1, base: {} }];
    expect(
      await conditionalViewportWithFormulas(rules, cells, async () => ({ error: '#VALUE!' })),
    ).toEqual([{}]);
    await expect(
      conditionalViewportWithFormulas(rules, cells, async () => ({ error: '#LIMIT!' })),
    ).rejects.toThrow('calculation exceeded');
    await expect(
      conditionalViewportWithFormulas(
        rules,
        [{ ...cells[0], value: { error: '#LIMIT!' } }],
        async () => true,
      ),
    ).rejects.toThrow('calculation exceeded');
    const controller = new AbortController();
    let started!: () => void;
    const ready = new Promise<void>((resolve) => {
      started = resolve;
    });
    const pending = conditionalViewportWithFormulas(
      rules,
      cells,
      () => {
        started();
        return new Promise(() => {});
      },
      controller.signal,
    );
    await ready;
    controller.abort();
    await expect(pending).rejects.toMatchObject({ name: 'AbortError' });
  });

  it('keeps rule and cell inputs stable while awaiting a formula result', async () => {
    const rules = validateConditionalRules(
      [{ ...numberRule('f'), condition: { kind: 'formula', expression: 'A1>0' } }],
      10,
      10,
    );
    const cells = [{ row: 1, column: 2, value: 1, base: { background: '#ffffff' } }];
    let resolve!: (value: boolean) => void;
    let started!: () => void;
    const ready = new Promise<void>((done) => {
      started = done;
    });
    const pending = conditionalViewportWithFormulas(rules, cells, () => {
      started();
      return new Promise<boolean>((done) => {
        resolve = done;
      });
    });
    await ready;
    rules[0].style.color = '#0000ff';
    cells[0].base.background = '#000000';
    resolve(true);
    expect(await pending).toEqual([{ background: '#ffffff', color: '#ff0000', bold: true }]);
  });

  it('rejects oversized viewport input before cloning it', async () => {
    let baseReads = 0;
    const cell = {
      row: 1,
      column: 2,
      value: 'a'.repeat(32768),
      get base() {
        baseReads++;
        return {};
      },
    };
    await expect(
      conditionalViewportWithFormulas(
        [],
        Array.from({ length: 400 }, () => cell),
        async () => false,
      ),
    ).rejects.toThrow('Conditional formatting viewport exceeds text work budget');
    expect(baseReads).toBe(0);
  });

  it('rewrites relative and explicit cross-sheet formula references on structural changes', () => {
    const rules = [{ ...numberRule('f'), condition: { kind: 'formula', expression: 'A2>$B$1' } }];
    const inserted = transformConditionalRules(rules, 10, 10, 'row', 'insert', 0, 1, 'Data');
    expect(inserted[0].range.startRow).toBe(2);
    expect(inserted[0].condition).toEqual({ kind: 'formula', expression: '(A3>$B$2)' });
    expect(() => transformConditionalRules(rules, 10, 10, 'row', 'insert', 0, 1)).toThrow();
    const external = [
      { ...numberRule('f'), condition: { kind: 'formula', expression: 'Other!A1>0' } },
    ];
    const changed = transformConditionalRules(external, 10, 10, 'row', 'insert', 0, 1, 'Data', {
      name: 'Other',
      axisLength: 10,
    });
    expect(changed[0].range).toEqual(range);
    expect(changed[0].condition).toEqual({ kind: 'formula', expression: "('Other'!A2>0)" });
  });

  it('transforms ranges on insertion and deletion without changing priority', () => {
    const rules = [
      numberRule('first'),
      { ...numberRule('second'), range: { startRow: 7, endRow: 9, startColumn: 2, endColumn: 6 } },
    ];
    const inserted = transformConditionalRules(rules, 10, 10, 'row', 'insert', 3, 2);
    expect(inserted.map((rule) => [rule.id, rule.range.startRow, rule.range.endRow])).toEqual([
      ['first', 1, 7],
      ['second', 9, 11],
    ]);
    const removed = transformConditionalRules(inserted, 12, 10, 'row', 'delete', 1, 6);
    expect(removed.map((rule) => [rule.id, rule.range.startRow, rule.range.endRow])).toEqual([
      ['second', 3, 5],
    ]);
    expect(() => transformConditionalRules([], 10, 10, 'row', 'delete', 0, 10)).toThrow();
  });

  it('matches an independent numeric oracle across generated coordinates and thresholds', () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 0, max: 9 }),
        fc.integer({ min: 0, max: 9 }),
        fc.integer({ min: -100, max: 100 }),
        fc.integer({ min: -100, max: 100 }),
        (row, column, value, threshold) => {
          const rules = validateConditionalRules([numberRule('x', threshold)], 10, 10);
          const actual = conditionalStyle(rules, row, column, value, {});
          const expected =
            row >= 1 && row < 5 && column >= 2 && column < 6 && value >= threshold
              ? { color: '#ff0000', bold: true }
              : {};
          expect(actual).toEqual(expected);
        },
      ),
      { seed: 20261112, numRuns: 3000 },
    );
  });
});
