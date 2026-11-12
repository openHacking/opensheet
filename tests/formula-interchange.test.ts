import { expect, it } from 'vitest';
import { evaluateFormula, normalizeFormulaPrecedence, parseFormula } from '@opensheetjs/formula';
import { parseSpreadsheetFormula } from '../packages/formula/src/parser.js';
import {
  createPrecedenceNormalizer,
  mayDifferInPowerPrecedence,
} from '../packages/adapter-sheetjs/src/formula.js';
import fc from 'fast-check';

it.each([
  ['2^3^2', 64, 512],
  ['-2^2', 4, -4],
  ['2^(-3^2)', 512, 1 / 512],
  ['2^-2^2', 1 / 16, 1 / 16],
  ['ROUND(2^3^2,0)', 64, 512],
])('preserves each precedence profile for %s', (expression, spreadsheet, native) => {
  const imported = normalizeFormulaPrecedence(expression, 'spreadsheet');
  const exported = normalizeFormulaPrecedence(expression, 'native');
  expect(imported).toBeDefined();
  expect(exported).toBeDefined();
  expect(evaluateFormula(parseFormula(imported!), () => null)).toBe(spreadsheet);
  expect(evaluateFormula(parseSpreadsheetFormula(exported!), () => null)).toBe(native);
});

it('preserves supported quoted references and declines unsafe rewrites', () => {
  const input = "'Other Sheet'!$A1^2^2";
  const normalized = normalizeFormulaPrecedence(input, 'spreadsheet');
  expect(normalized).toBeDefined();
  expect(evaluateFormula(parseFormula(normalized!), () => 2)).toBe(16);
  expect(normalizeFormulaPrecedence('UNKNOWN(1)^-2', 'spreadsheet')).toBeUndefined();
  expect(normalizeFormulaPrecedence('1e309^2^2', 'spreadsheet')).toBeUndefined();
  expect(normalizeFormulaPrecedence('A1^', 'spreadsheet')).toBeUndefined();
});

it('keeps conversion profiles distinct after the bounded memo fills', () => {
  const normalize = createPrecedenceNormalizer();
  for (let index = 0; index < 400; index++)
    expect(normalize(`2^3^2+${index}`, 'spreadsheet')).toBe(`(((2^3)^2)+${index})`);
  expect(normalize('2^3^2', 'native')).toBe('(2^(3^2))');
  expect(normalize('2^3^2', 'spreadsheet')).toBe('((2^3)^2)');
  expect(normalize('UNKNOWN(1)^-2', 'spreadsheet')).toBeUndefined();
});

it('detects generated expressions whose precedence trees differ', () => {
  const atom = fc.constantFrom('1', '2', '3e-2', 'A1', "'A^B'!A1", '"-^"');
  let expression: fc.Arbitrary<string> = atom;
  for (let depth = 0; depth < 3; depth++) {
    const prior = expression;
    expression = fc.oneof(
      atom,
      prior.map((value) => `(${value})`),
      prior.map((value) => `-${value}`),
      prior.map((value) => `+${value}`),
      prior.map((value) => `${value}%`),
      fc
        .tuple(prior, fc.constantFrom('^', '+', '-', '*', '/'), prior)
        .map((parts) => parts.join('')),
    );
  }
  fc.assert(
    fc.property(expression, (formula) => {
      const native = parseFormula(formula);
      const spreadsheet = parseSpreadsheetFormula(formula);
      if (JSON.stringify(native) !== JSON.stringify(spreadsheet))
        expect(mayDifferInPowerPrecedence(formula), formula).toBe(true);
    }),
    { seed: 20261111, numRuns: 10_000 },
  );
});
