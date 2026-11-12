import { expect, it } from 'vitest';
import { parseFormula } from '@opensheetjs/formula';
import {
  rewriteValidationSheetReferences,
  transformValidationEntries,
  validateValidationEntries,
} from '../packages/core/src/validation-model.js';

it('rewrites explicit validation sheet references on rename and invalidates deleted targets', () => {
  const source = [
    {
      ...entry('f', 1, 4),
      rule: { kind: 'formula', expression: 'lower!A1>A1', allowBlank: false, policy: 'warn' },
    },
  ];
  const renamed = rewriteValidationSheetReferences(source, 10, 4, 'Lower', 'New name');
  const removed = rewriteValidationSheetReferences(renamed, 10, 4, 'NEW NAME', null);
  expect(renamed[0].rule).toMatchObject({ expression: "('New name'!A1>A1)" });
  expect(removed[0].rule).toMatchObject({ expression: '(#REF!>A1)' });
  expect(source[0].rule.expression).toBe('lower!A1>A1');
});

it('keeps shifted validation formulas parseable at the A1 address boundary', () => {
  const source = [
    {
      id: 'limit',
      range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
      rule: {
        kind: 'formula',
        expression: 'A1048576>0',
        allowBlank: false,
        policy: 'reject',
      },
    },
  ];
  const result = transformValidationEntries(source, 5, 4, 'row', 'insert', 0, 1, 'Data');
  expect(result[0].rule).toMatchObject({ expression: '(#REF!>0)' });
  expect(() => parseFormula((result[0].rule as { expression: string }).expression)).not.toThrow();
});

const rule = { kind: 'number', operator: 'gte', first: 1, allowBlank: false, policy: 'reject' };
const entry = (id: string, startRow: number, endRow: number) => ({
  id,
  range: { startRow, endRow, startColumn: 0, endColumn: 2 },
  rule,
});

it('validates identities and sheet bounds while preserving independent overlapping rules', () => {
  const source = [entry('first', 1, 4), entry('second', 2, 5)];
  const result = validateValidationEntries(source, 10, 4);
  expect(result).toEqual(source);
  result[0].range.startRow = 0;
  expect(source[0].range.startRow).toBe(1);
  expect(() => validateValidationEntries([source[0], source[0]], 10, 4)).toThrow('identity');
  expect(() => validateValidationEntries([entry('out', 1, 11)], 10, 4)).toThrow('range');
  expect(() => validateValidationEntries([{ ...source[0], unknown: true }], 10, 4)).toThrow(
    'entry',
  );
  expect(() =>
    validateValidationEntries(
      [{ ...source[0], range: { ...source[0].range, unknown: true } }],
      10,
      4,
    ),
  ).toThrow('range');
});

it('rejects malformed validation formulas before accepting the rule collection', () => {
  const source = {
    ...entry('formula', 0, 1),
    rule: { kind: 'formula', expression: 'A1+', allowBlank: false, policy: 'reject' },
  };
  expect(() => validateValidationEntries([source], 10, 4)).toThrow(
    expect.objectContaining({
      code: 'INVALID_ARGUMENT',
      message: 'Invalid validation formula syntax',
    }),
  );
  expect(source.rule.expression).toBe('A1+');
  for (const expression of ['A1>0', '#REF!', 'UNKNOWN(A1)']) {
    expect(
      validateValidationEntries([{ ...source, rule: { ...source.rule, expression } }], 10, 4)[0]
        .rule,
    ).toMatchObject({ expression });
  }
});

it('transforms all rules across structural edits and removes fully deleted ranges', () => {
  const source = [entry('first', 1, 4), entry('second', 3, 5), entry('deleted', 4, 5)];
  const inserted = transformValidationEntries(source, 10, 4, 'row', 'insert', 3, 2);
  expect(inserted.map((item) => [item.id, item.range.startRow, item.range.endRow])).toEqual([
    ['first', 1, 6],
    ['second', 5, 7],
    ['deleted', 6, 7],
  ]);
  const removed = transformValidationEntries(inserted, 12, 4, 'row', 'delete', 3, 4);
  expect(removed.map((item) => [item.id, item.range.startRow, item.range.endRow])).toEqual([
    ['first', 1, 3],
  ]);
  expect(source.map((item) => item.range.startRow)).toEqual([1, 3, 4]);
  expect(() => transformValidationEntries([], 10, 4, 'row', 'delete', 9, 2)).toThrow(
    'transformation',
  );
});

it('bounds collection-wide list text before cloning', () => {
  const list = { kind: 'list', values: ['x'.repeat(32768)], allowBlank: true, policy: 'warn' };
  const entries = Array.from({ length: 31 }, (_, i) => ({ ...entry(String(i), 0, 1), rule: list }));
  expect(validateValidationEntries(entries.slice(0, 30), 10, 4)).toHaveLength(30);
  expect(() => validateValidationEntries(entries, 10, 4)).toThrow('work budget');
});

it('rewrites local references and re-anchors relative formulas after origin deletion', () => {
  const formula = {
    ...entry('formula', 1, 4),
    rule: {
      kind: 'formula',
      expression: 'AND(A2>Other!A2,$A$2>0,Data!A2>0)',
      allowBlank: false,
      policy: 'reject',
    },
  };
  const expression = (result: ReturnType<typeof transformValidationEntries>) => {
    const rule = result[0].rule;
    if (rule.kind !== 'formula') throw new Error('Expected formula');
    return parseFormula(rule.expression);
  };
  const inserted = transformValidationEntries([formula], 10, 4, 'row', 'insert', 1, 1, 'Data');
  expect(inserted[0].range.startRow).toBe(2);
  expect(expression(inserted)).toEqual(parseFormula('AND(A3>Other!A2,$A$3>0,Data!A3>0)'));
  const removed = transformValidationEntries([formula], 10, 4, 'row', 'delete', 1, 1, 'Data');
  expect(removed[0].range).toMatchObject({ startRow: 1, endRow: 3 });
  expect(expression(removed)).toEqual(parseFormula('AND(A2>Other!A3,#REF!>0,Data!A2>0)'));
  expect(formula.rule.expression).toBe('AND(A2>Other!A2,$A$2>0,Data!A2>0)');
});

it('keeps the owner range and unrelated references fixed during an edit on another sheet', () => {
  const formula = {
    ...entry('formula', 1, 4),
    rule: {
      kind: 'formula',
      expression: 'AND(A2>Limits!$B$2,Other!B2>0)',
      allowBlank: false,
      policy: 'reject',
    },
  };
  const inserted = transformValidationEntries([formula], 10, 2, 'column', 'insert', 1, 2, 'Data', {
    name: 'limits',
    axisLength: 5,
  });
  expect(inserted[0].range).toEqual(formula.range);
  const rule = inserted[0].rule;
  if (rule.kind !== 'formula') throw new Error('Expected formula');
  expect(parseFormula(rule.expression)).toEqual(parseFormula('AND(A2>Limits!$D$2,Other!B2>0)'));
  const removed = transformValidationEntries([formula], 10, 2, 'column', 'delete', 1, 1, 'Data', {
    name: 'Limits',
    axisLength: 5,
  });
  const deletedRule = removed[0].rule;
  if (deletedRule.kind !== 'formula') throw new Error('Expected formula');
  expect(parseFormula(deletedRule.expression)).toEqual(parseFormula('AND(A2>#REF!,Other!B2>0)'));
  expect(removed[0].range).toEqual(formula.range);
  expect(
    transformValidationEntries([formula], 10, 2, 'column', 'insert', 1, 2, 'Data', {
      name: 'Unused',
      axisLength: 5,
    }),
  ).toEqual([formula]);
});

it('shrinks formula ranges after partial deletion and preserves fully deleted reference errors', () => {
  const formula = {
    ...entry('formula', 6, 8),
    rule: {
      kind: 'formula',
      expression: 'AND(SUM(A1:A5)>0,SUM($A$5:$A$2)>0,SUM(A1:A2)>0,SUM(Other!A1:A5)>0)',
      allowBlank: false,
      policy: 'reject',
    },
  };
  const result = transformValidationEntries([formula], 10, 4, 'row', 'delete', 0, 2, 'Data');
  const rule = result[0].rule;
  if (rule.kind !== 'formula') throw new Error('Expected formula');
  expect(parseFormula(rule.expression)).toEqual(
    parseFormula('AND(SUM(A1:A3)>0,SUM($A$3:$A$1)>0,SUM(#REF!)>0,SUM(Other!A1:A5)>0)'),
  );
  expect(result[0].range).toMatchObject({ startRow: 4, endRow: 6 });
  expect(formula.rule.expression).toContain('A1:A5');

  const cross = { ...formula, rule: { ...formula.rule, expression: 'SUM(Limits!A1:C1)>0' } };
  const other = transformValidationEntries([cross], 10, 4, 'column', 'delete', 0, 1, 'Data', {
    name: 'Limits',
    axisLength: 4,
  });
  const crossRule = other[0].rule;
  if (crossRule.kind !== 'formula') throw new Error('Expected formula');
  expect(parseFormula(crossRule.expression)).toEqual(parseFormula('SUM(Limits!A1:B1)>0'));
  expect(other[0].range).toEqual(cross.range);

  const anchored = {
    ...entry('anchored', 1, 4),
    rule: { ...formula.rule, expression: 'AND(SUM(A2:A5)>0,SUM($A$2:$A$5)>0)' },
  };
  const moved = transformValidationEntries([anchored], 10, 4, 'row', 'delete', 1, 1, 'Data');
  const movedRule = moved[0].rule;
  if (movedRule.kind !== 'formula') throw new Error('Expected formula');
  expect(parseFormula(movedRule.expression)).toEqual(
    parseFormula('AND(SUM(A2:A5)>0,SUM($A$2:$A$4)>0)'),
  );
});
