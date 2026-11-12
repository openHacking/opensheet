import { expect, it } from 'vitest';
import { sample } from '../apps/playground/src/budget.js';
import { FileReader } from '@opensheetjs/core';

it('provides independent validated budget snapshots with formulas and formatting', async () => {
  const first = await sample();
  const second = await sample();
  first.sheets[0].name = 'Changed';
  expect(second.sheets[0].name).toBe('Launch budget');
  const reader = new FileReader(second);
  try {
    const sheet = reader.getSheetById(second.sheets[0].id)!;
    expect(sheet.range('A2').getValues()).toEqual([['Brand identity']]);
    expect(sheet.range('F11').getValues()).toEqual([[13097]]);
    expect(sheet.range('F11').getDisplayValues()).toEqual([['$13,097.00']]);
  } finally {
    reader.dispose();
  }
});
