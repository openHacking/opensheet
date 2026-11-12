import { describe, expect, it, vi } from 'vitest';
import fc from 'fast-check';
import {
  indexNamedRanges,
  removeNamedRangeSheet,
  transformNamedRanges,
  transformStagedNamedRanges,
  validateNamedRanges,
  validateStagedNamedRanges,
  type NamedRange,
} from '../packages/core/src/named-range-model.js';

const sheets = [
  { id: 'first', rows: 10, columns: 8 },
  { id: 'second', rows: 10, columns: 8 },
];
const target = {
  sheetId: 'first',
  range: { startRow: 2, startColumn: 1, endRow: 5, endColumn: 4 },
};
const names: NamedRange[] = [
  { id: 'global', name: 'Revenue', scopeSheetId: null, target },
  { id: 'local', name: 'revenue', scopeSheetId: 'second', target: null },
  {
    id: 'other',
    name: 'Costs',
    scopeSheetId: 'first',
    target: { sheetId: 'second', range: { startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 } },
  },
];

describe('internal named-range model', () => {
  it('validates and transforms a stable staged revision without publishing partial results', async () => {
    const input: NamedRange[] = Array.from({ length: 8192 }, (_, i) => ({
      id: `name_${i}`,
      name: `Name_${i}`,
      scopeSheetId: null,
      target,
    }));
    const original = structuredClone(input);
    expect(await validateStagedNamedRanges(input, sheets)).toEqual(
      validateNamedRanges(input, sheets),
    );
    expect(await transformStagedNamedRanges(input, sheets, 'first', 'row', 'insert', 3, 1)).toEqual(
      transformNamedRanges(input, sheets, 'first', 'row', 'insert', 3, 1),
    );
    expect(input).toEqual(original);
    const controller = new AbortController();
    const reason = new Error('cancelled');
    let clock = 0;
    const now = vi.spyOn(performance, 'now').mockImplementation(() => (clock += 10));
    try {
      const pending = transformStagedNamedRanges(
        input,
        sheets,
        'first',
        'row',
        'insert',
        3,
        1,
        controller.signal,
      );
      setTimeout(() => controller.abort(reason), 0);
      await expect(pending).rejects.toBe(reason);
    } finally {
      now.mockRestore();
    }
    expect(input).toEqual(original);
  });
  it('uses local shadowing, retains broken names and detaches inputs and lookup results', () => {
    const input = structuredClone(names),
      index = indexNamedRanges(input, sheets);
    input[0].name = 'Changed';
    expect(index.resolve('REVENUE', 'first')).toEqual(names[0]);
    expect(index.resolve('Revenue', 'second')).toEqual(names[1]);
    expect(index.resolve('Missing', 'first')).toBeUndefined();
    const result = index.resolve('Revenue', 'first')!;
    result.target!.range.startRow = 9;
    expect(index.resolve('Revenue', 'first')).toEqual(names[0]);
    expect(() => index.resolve('Revenue', 'missing')).toThrow('Invalid named-range lookup');
  });

  it('rejects ambiguous names, duplicates, invalid scopes and malformed targets', () => {
    for (const name of [
      '',
      'A1',
      'A1foo',
      'XFE123',
      'R',
      'c',
      'TRUE',
      '9bad',
      'a b',
      '中',
      'x'.repeat(256),
    ])
      expect(() => validateNamedRanges([{ ...names[0], name }], sheets), name).toThrow();
    for (const name of ['A0', '_name', 'Revenue.Q1', 'sum'])
      expect(validateNamedRanges([{ ...names[0], name }], sheets)[0].name).toBe(name);
    expect(() =>
      validateNamedRanges([names[0], { ...names[0], id: 'next', name: 'REVENUE' }], sheets),
    ).toThrow('Duplicate');
    expect(() => validateNamedRanges([names[0], { ...names[1], id: 'global' }], sheets)).toThrow(
      'identity',
    );
    expect(() => validateNamedRanges([{ ...names[0], scopeSheetId: 'missing' }], sheets)).toThrow(
      'scope',
    );
    expect(() =>
      validateNamedRanges(
        [{ ...names[0], target: { ...target, range: { ...target.range, endRow: 11 } } }],
        sheets,
      ),
    ).toThrow('target');
    expect(() => validateNamedRanges([{ ...names[0], extra: true }], sheets)).toThrow(
      'Invalid named range',
    );
    expect(() =>
      validateNamedRanges(
        [names[0], { ...names[2], target: { ...target, range: { ...target.range, extra: 1 } } }],
        sheets,
      ),
    ).toThrow('target');
    expect(validateNamedRanges(names, sheets)).toEqual(names);
    expect(() => validateNamedRanges(new Array(1), sheets)).toThrow('Invalid named range');
    expect(() => validateNamedRanges([Object.create(names[0])], sheets)).toThrow(
      'Invalid named range',
    );
    expect(() =>
      validateNamedRanges([{ ...names[0], target: Object.create(target) }], sheets),
    ).toThrow('target');
    expect(() =>
      validateNamedRanges(
        [{ ...names[0], target: { sheetId: 'first', range: Object.create(target.range) } }],
        sheets,
      ),
    ).toThrow('target');
    expect(() => validateNamedRanges([], [Object.create(sheets[0])])).toThrow('sheet');
    expect(() =>
      validateNamedRanges(
        Array.from({ length: 8193 }, () => names[0]),
        sheets,
      ),
    ).toThrow('Too many');
  });

  it('shifts, expands and breaks targets without mutating existing metadata', () => {
    const original = structuredClone(names);
    const inserted = transformNamedRanges(names, sheets, 'first', 'column', 'insert', 2, 2);
    expect(inserted[0].target?.range).toEqual({
      startRow: 2,
      startColumn: 1,
      endRow: 5,
      endColumn: 6,
    });
    expect(inserted[1]).toEqual(names[1]);
    expect(inserted[2]).toEqual(names[2]);
    const deleted = transformNamedRanges(names, sheets, 'first', 'row', 'delete', 2, 3);
    expect(deleted[0].target).toBeNull();
    expect(
      indexNamedRanges(deleted, [{ ...sheets[0], rows: 7 }, sheets[1]]).resolve('revenue', 'first')
        ?.id,
    ).toBe('global');
    expect(names).toEqual(original);
    expect(() => transformNamedRanges(names, sheets, 'first', 'row', 'delete', 0, 10)).toThrow();
    expect(names).toEqual(original);
  });

  it('removes sheet-scoped names and breaks surviving targets when a sheet is deleted', () => {
    const result = removeNamedRangeSheet(names, sheets, 'first');
    expect(result).toEqual([{ ...names[0], target: null }, names[1]]);
    expect(names[0].target).toEqual(target);
    expect(() => removeNamedRangeSheet(names, sheets, 'missing')).toThrow();
    expect(() => removeNamedRangeSheet([], [sheets[0]], 'first')).toThrow();
  });

  it('matches an independent seeded membership model across row and column edits', () => {
    fc.assert(
      fc.property(
        fc.integer({ min: 2, max: 40 }),
        fc.nat(),
        fc.nat(),
        fc.nat(),
        fc.nat(),
        fc.boolean(),
        fc.boolean(),
        (length, a, b, c, d, row, insert) => {
          const start = a % length,
            end = start + 1 + (b % (length - start));
          const position = c % (insert ? length + 1 : length);
          const count = insert ? 1 + (d % 4) : 1 + (d % Math.min(length - position, length - 1));
          const expected = Array.from({ length }, (_, i) => i >= start && i < end);
          if (insert)
            expected.splice(position, 0, ...Array(count).fill(position > start && position < end));
          else expected.splice(position, count);
          const sheet = { id: 'sheet', rows: row ? length : 1, columns: row ? 1 : length };
          const range = row
            ? { startRow: start, endRow: end, startColumn: 0, endColumn: 1 }
            : { startRow: 0, endRow: 1, startColumn: start, endColumn: end };
          const entry = {
            id: 'range',
            name: 'Values',
            scopeSheetId: null,
            target: { sheetId: 'sheet', range },
          };
          const actual = transformNamedRanges(
            [entry],
            [sheet],
            'sheet',
            row ? 'row' : 'column',
            insert ? 'insert' : 'delete',
            position,
            count,
          )[0].target?.range;
          expect(
            expected.map(
              (_, i) =>
                !!actual &&
                i >= (row ? actual.startRow : actual.startColumn) &&
                i < (row ? actual.endRow : actual.endColumn),
            ),
          ).toEqual(expected);
        },
      ),
      { seed: 20261112, numRuns: 2000 },
    );
  });
});
