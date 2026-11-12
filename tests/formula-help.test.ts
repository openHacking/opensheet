import { describe, expect, it } from 'vitest';
import { formulaHelp, functionDefinitions } from '@opensheetjs/formula';

describe('incomplete formula help', () => {
  it('completes a token at the cursor using the evaluator registry', () => {
    const help = formulaHelp('=suMM(1)', 3);
    expect(help.completion?.start).toBe(1);
    expect(help.completion?.end).toBe(5);
    expect(help.completion?.functions.map((entry) => entry.name)).toEqual([
      'SUM',
      'SUMIF',
      'SUMIFS',
    ]);
    expect(help.completion?.functions[0]).toBe(
      functionDefinitions.find((entry) => entry.name === 'SUM'),
    );
  });
  it('tracks nested calls and ignores quoted commas and escaped quotes', () => {
    expect(formulaHelp('=IF(SUM(1,2),"a,""b",').call).toMatchObject({
      definition: { name: 'IF' },
      argument: 2,
    });
    expect(formulaHelp('=SUM(1,ROUND(2,').call).toMatchObject({
      definition: { name: 'ROUND' },
      argument: 1,
    });
    expect(formulaHelp("=SUM('Sheet, one'!A1,").call?.argument).toBe(1);
    expect(formulaHelp('=SUM((1+2),').call?.argument).toBe(1);
  });
  it('does not suggest functions inside text or qualified sheet references', () => {
    for (const input of ['SUM', '=CONCAT("SU', "='SU", '=Sheet!SU', '=42SU'])
      expect(formulaHelp(input).completion).toBeUndefined();
    expect(formulaHelp('=SUM(1)').call).toBeUndefined();
  });
  it('bounds input and rejects invalid cursor offsets without evaluating', () => {
    expect(formulaHelp('=' + 'S'.repeat(32767))).toEqual({});
    for (const offset of [-1, 0, 20, NaN, 1.5]) expect(formulaHelp('=SUM', offset)).toEqual({});
    expect(formulaHelp('=IF(1/0,SU').completion?.functions[0]?.name).toBe('SUM');
  });
});
