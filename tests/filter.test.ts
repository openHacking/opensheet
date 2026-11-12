import { describe, expect, it } from 'vitest';
import { createWorkbook, openWorkbook, type FilterCriterion } from '@opensheetjs/core';
import { matchesFilter } from '../packages/core/src/filter.js';

describe('typed multi-column filters', () => {
  it('distinguishes numbers, text, booleans, empty values and formula errors', () => {
    const cases: Array<[unknown, FilterCriterion, boolean]> = [
      [2, { column: 0, operator: 'eq', value: 2 }, true],
      ['2', { column: 0, operator: 'eq', value: 2 }, false],
      [true, { column: 0, operator: 'eq', value: 1 }, false],
      ['ALPHA', { column: 0, operator: 'eq', value: 'alpha' }, true],
      ['2', { column: 0, operator: 'gt', value: 1 }, false],
      [3, { column: 0, operator: 'gte', value: 3 }, true],
      [3, { column: 0, operator: 'lt', value: 3 }, false],
      [3, { column: 0, operator: 'lte', value: 3 }, true],
      ['Alpha', { column: 0, operator: 'startsWith', value: 'AL' }, true],
      ['Alpha', { column: 0, operator: 'endsWith', value: 'HA' }, true],
      ['Alpha', { column: 0, operator: 'contains', value: 'PH' }, true],
      ['Alpha', { column: 0, operator: 'notContains', value: 'beta' }, true],
      [2, { column: 0, operator: 'notContains', value: 'a' }, false],
      [null, { column: 0, operator: 'neq', value: 1 }, false],
      [{ error: '#REF!' }, { column: 0, operator: 'neq', value: 1 }, false],
      [{ error: '#REF!' }, { column: 0, operator: 'isError' }, true],
      ['', { column: 0, operator: 'isBlank' }, true],
      [null, { column: 0, operator: 'isBlank' }, true],
      [' ', { column: 0, operator: 'isNotBlank' }, true],
      [null, { column: 0, operator: 'isNotError' }, true],
    ];
    for (const [value, criterion, expected] of cases)
      expect(matchesFilter(value as Parameters<typeof matchesFilter>[0], criterion)).toBe(expected);
  });

  it('recalculates formula criteria from staged data and preserves history/reopen', async () => {
    let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 5 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:B5').setValues([
        ['Group', 'Value'],
        ['a', 2],
        ['b', 5],
        ['A', 4],
        ['a', '5'],
      ]);
      await sheet.range('C2:C5').setFormulas([['B2*2'], ['B3*2'], ['B4*2'], ['B5*2']]);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 1,
        criteria: [
          { column: 0, operator: 'eq', value: 'a' },
          { column: 2, operator: 'gt', value: 5 },
          { column: 1, operator: 'lt', value: 5 },
        ],
      });
      expect(book.filteredRows(sheet.id)).toEqual([1, 2, 4]);
      await book.transaction({}, async (tx) => {
        await tx
          .getSheets()[0]
          .range('B2')
          .setValues([[3]]);
      });
      expect(book.filteredRows(sheet.id)).toEqual([2, 4]);
      await book.undo();
      expect(book.filteredRows(sheet.id)).toEqual([1, 2, 4]);
      await book.redo();
      const id = book.id;
      await book.close();
      book = await openWorkbook(id, { cacheBytes: 0 });
      expect(book.filteredRows(sheet.id)).toEqual([2, 4]);
      const reopened = book.getSheets()[0];
      await reopened.setFilter(null);
      expect(book.filteredRows(sheet.id)).toEqual([]);
      await book.undo();
      expect(book.filteredRows(sheet.id)).toEqual([2, 4]);
    } finally {
      await book.close();
    }
  });

  it('supports any criteria, zero headers and formula errors explicitly', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 4 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A4').setValues([[false], [true], ['true'], [null]]);
      await sheet.range('B4').setFormulas([['1/0']]);
      await sheet.setFilter({
        mode: 'any',
        headerRows: 0,
        criteria: [
          { column: 0, operator: 'eq', value: true },
          { column: 1, operator: 'isError' },
        ],
      });
      expect(book.filteredRows(sheet.id)).toEqual([0, 2]);
    } finally {
      await book.close();
    }
  });

  it('matches folded Unicode and empty text queries in staged filters', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 4 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A4').setValues([['İSTANBUL'], ['istanbul'], [''], [42]]);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 0, operator: 'startsWith', value: 'İS' }],
      });
      expect(book.filteredRows(sheet.id)).toEqual([1, 2, 3]);
      await sheet.setFilter({
        mode: 'any',
        headerRows: 0,
        criteria: [{ column: 0, operator: 'contains', value: '' }],
      });
      expect(book.filteredRows(sheet.id)).toEqual([3]);
      await sheet.setFilter({ column: 0, query: 'İS' });
      expect(book.filteredRows(sheet.id)).toEqual([1, 2, 3]);
    } finally {
      await book.close();
    }
  });

  it('shifts criteria and headers with structural edits and restores removed criteria on undo', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 4, columns: 4 }] });
    try {
      const sheet = book.getSheets()[0];
      const criteria: FilterCriterion[] = [
        { column: 0, operator: 'isBlank' },
        { column: 2, operator: 'isBlank' },
      ];
      await sheet.setFilter({ mode: 'all', headerRows: 1, criteria });
      await sheet.insertColumns(1);
      expect(book.sheetData(sheet.id).filter).toMatchObject({
        criteria: [criteria[0], { ...criteria[1], column: 3 }],
      });
      await sheet.deleteColumns(0);
      expect(book.sheetData(sheet.id).filter).toMatchObject({ criteria: [criteria[1]] });
      await sheet.deleteColumns(2);
      expect(book.sheetData(sheet.id).filter).toBeNull();
      await book.undo();
      await sheet.insertRows(0, 2);
      expect(book.sheetData(sheet.id).filter).toMatchObject({ headerRows: 3 });
      await sheet.deleteRows(1, 3);
      expect(book.sheetData(sheet.id).filter).toMatchObject({ headerRows: 1 });
    } finally {
      await book.close();
    }
  });

  it('preserves legacy formatted substring matching and the first header row', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 3 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet.range('A1:A3').setValues([['Header'], ['Alpha'], ['beta']]);
      await sheet.setFilter({ column: 0, query: 'AL' });
      expect(book.filteredRows(sheet.id)).toEqual([2]);
      await sheet.insertColumns(0);
      expect(book.sheetData(sheet.id).filter).toEqual({ column: 1, query: 'AL' });
    } finally {
      await book.close();
    }
  });

  it('preserves typed criteria through native JSON export and import', async () => {
    const source = await createWorkbook({ sheets: [{ name: 'Data', rows: 3 }] });
    const target = await createWorkbook();
    try {
      const sheet = source.getSheets()[0];
      await sheet.range('A1:A3').setValues([[1], [2], ['3']]);
      await sheet.setFilter({
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 0, operator: 'gt', value: 1 }],
      });
      const file = JSON.parse(JSON.stringify(await source.toJSON()));
      await target.importJSON(file);
      const restored = target.getSheets()[0];
      expect(target.sheetData(restored.id).filter).toEqual(source.sheetData(sheet.id).filter);
      expect(target.filteredRows(restored.id)).toEqual([0, 2]);
    } finally {
      await source.close();
      await target.close();
    }
  });

  it('aborts exhausted formula evaluation without committing a filter or visibility', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 150 }] });
    try {
      const sheet = book.getSheets()[0];
      await sheet
        .range('B1:B140')
        .setFormulas(Array.from({ length: 140 }, (_, i) => [i === 139 ? '1' : `B${i + 2}`]));
      await sheet.range('A1').setFormulas([['B1']]);
      const revision = book.revision;
      await expect(
        sheet.setFilter({
          mode: 'all',
          headerRows: 0,
          criteria: [{ column: 0, operator: 'isError' }],
        }),
      ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
      expect(book.revision).toBe(revision);
      expect(book.sheetData(sheet.id).filter).toBeNull();
      expect(book.filteredRows(sheet.id)).toEqual([]);
    } finally {
      await book.close();
    }
  });

  it('rejects invalid filters atomically', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
    try {
      const sheet = book.getSheets()[0],
        revision = book.revision;
      for (const filter of [
        { mode: 'all', headerRows: 0, criteria: [] },
        { mode: 'all', headerRows: 3, criteria: [{ column: 0, operator: 'isBlank' }] },
        { mode: 'all', headerRows: 0, criteria: [{ column: 2, operator: 'isBlank' }] },
        { mode: 'all', headerRows: 0, criteria: [{ column: 0, operator: 'gt', value: '2' }] },
        {
          column: 0,
          query: 'x',
          mode: 'all',
          headerRows: 0,
          criteria: [{ column: 0, operator: 'isBlank' }],
        },
      ]) {
        await expect(
          book.execute({ type: 'core.sheet.filter', payload: { sheetId: sheet.id, filter } }),
        ).rejects.toThrow();
        expect(book.revision).toBe(revision);
        expect(book.sheetData(sheet.id).filter).toBeNull();
      }
    } finally {
      await book.close();
    }
  });
});
