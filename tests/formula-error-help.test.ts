import { expect, it } from 'vitest';
import { formulaErrorExplanation } from '@opensheetjs/core';

it('provides general guidance for supported calculation errors only', () => {
  for (const code of [
    '#DIV/0!',
    '#VALUE!',
    '#REF!',
    '#NAME?',
    '#N/A',
    '#NUM!',
    '#CYCLE!',
    '#LIMIT!',
  ])
    expect(formulaErrorExplanation(code)).toBeTruthy();
  expect(formulaErrorExplanation('#LOADING')).toBeUndefined();
  expect(formulaErrorExplanation('toString')).toBeUndefined();
});
