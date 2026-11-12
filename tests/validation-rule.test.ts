import { describe, expect, it } from 'vitest';
import {
  matchesValidationRule,
  validationRuleSchema,
} from '../packages/core/src/validation-rule.js';

const base = { allowBlank: false, policy: 'reject' } as const;
const rule = (value: unknown) => validationRuleSchema.parse(value);

describe('internal validation predicates', () => {
  it('validates range operands and rejects malformed or ambiguous rules', () => {
    for (const input of [
      { ...base, kind: 'number', operator: 'between', first: 1 },
      { ...base, kind: 'date', operator: 'between', first: 2, second: 1 },
      { ...base, kind: 'number', operator: 'eq', first: 1, second: 2 },
      { ...base, kind: 'textLength', operator: 'gt', first: -1 },
      { ...base, kind: 'textLength', operator: 'gt', first: 1.5 },
      { ...base, kind: 'list', values: [] },
      { ...base, kind: 'list', values: ['x'.repeat(32768), 'y'] },
      { ...base, kind: 'number', operator: 'eq', first: Infinity },
      { ...base, kind: 'number', operator: 'eq', first: 1, unexpected: true },
    ])
      expect(validationRuleSchema.safeParse(input).success).toBe(false);
  });
  it('checks inclusive numeric and date bounds without coercing strings or errors', () => {
    for (const kind of ['number', 'date']) {
      const inside = rule({ ...base, kind, operator: 'between', first: 1, second: 3 });
      const outside = rule({ ...base, kind, operator: 'notBetween', first: 1, second: 3 });
      for (const value of [1, 2, 3]) {
        expect(matchesValidationRule(inside, value)).toBe(true);
        expect(matchesValidationRule(outside, value)).toBe(false);
      }
      for (const value of [0, 4]) {
        expect(matchesValidationRule(inside, value)).toBe(false);
        expect(matchesValidationRule(outside, value)).toBe(true);
      }
      expect(matchesValidationRule(inside, '2')).toBe(false);
      expect(matchesValidationRule(inside, { error: '#VALUE!' })).toBe(false);
      expect(matchesValidationRule(inside, null)).toBe(false);
      expect(matchesValidationRule({ ...inside, allowBlank: true }, '')).toBe(true);
    }
  });
  it('implements every scalar comparison including exact boundaries', () => {
    for (const [operator, expected] of [
      ['eq', [false, true, false]],
      ['neq', [true, false, true]],
      ['lt', [true, false, false]],
      ['lte', [true, true, false]],
      ['gt', [false, false, true]],
      ['gte', [false, true, true]],
    ] as const) {
      const predicate = rule({ ...base, kind: 'number', operator, first: 2 });
      expect([1, 2, 3].map((value) => matchesValidationRule(predicate, value))).toEqual(expected);
    }
  });
  it('uses typed list members and Unicode code points for text length', () => {
    const list = rule({ ...base, kind: 'list', values: [1, '1', true] });
    expect(matchesValidationRule(list, 1)).toBe(true);
    expect(matchesValidationRule(list, '1')).toBe(true);
    expect(matchesValidationRule(list, false)).toBe(false);
    const text = rule({ ...base, kind: 'textLength', operator: 'eq', first: 2 });
    expect(matchesValidationRule(text, 'A😀')).toBe(true);
    expect(matchesValidationRule(text, '😀')).toBe(false);
    expect(matchesValidationRule(text, 12)).toBe(false);
  });
});
