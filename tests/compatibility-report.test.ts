import { describe, expect, it } from 'vitest';
import { add, finish, report } from '../packages/adapter-sheetjs/src/report.js';
import type { Issue } from '../packages/adapter-sheetjs/src/types.js';
import { createWorkbookFile, FileReader } from '@opensheetjs/core';
import { fromSheetJS, toSheetJS, writeXlsx } from '@opensheetjs/adapter-sheetjs';
import { createSnapshot } from '../packages/core/src/model.js';
import { encodeWorkbookFile } from '../packages/core/src/file-codec.js';
import * as XLSX from 'xlsx';

const issue = (overrides: Partial<Issue> = {}): Issue => ({
  code: 'FEATURE',
  severity: 'info',
  feature: 'test feature',
  action: 'preserved',
  message: 'Synthetic fixture',
  ...overrides,
});

describe('compatibility report safety', () => {
  it('exports the cell number format ahead of its style format', () => {
    const snapshot = createSnapshot();
    const sheet = snapshot.sheets[0];
    snapshot.styles.s1 = { numberFormat: '0.00' };
    sheet.cells['0:0'] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'number', value: 0.25 },
      styleId: 's1',
      numberFormat: '0%',
    };
    sheet.cells['1:0'] = {
      rowId: 1,
      columnId: 0,
      input: { type: 'number', value: 0.25 },
      styleId: 's1',
    };
    sheet.cells['2:0'] = {
      rowId: 2,
      columnId: 0,
      input: { type: 'blank' },
      styleId: 's1',
      link: { target: 'https://example.com/' },
    };
    const file = encodeWorkbookFile(snapshot);
    const { workbook } = toSheetJS(file);
    expect(workbook.Sheets[sheet.name].A1.z).toBe('0%');
    expect(workbook.Sheets[sheet.name].A2.z).toBe('0.00');
    expect(workbook.Sheets[sheet.name].A3.z).toBe('0.00');
    const reopened = XLSX.read(writeXlsx(file, XLSX).bytes, {
      type: 'array',
      cellNF: true,
      sheetStubs: true,
    });
    expect(reopened.Sheets[sheet.name].A1.z).toBe('0%');
    expect(reopened.Sheets[sheet.name].A2.z).toBe('0.00');
    expect(reopened.Sheets[sheet.name].A3.z).toBe('0.00');
    expect(reopened.Sheets[sheet.name].A3.l.Target).toBe('https://example.com/');
  });
  it('preserves blank-cell hyperlinks without changing blank values in XLSX', () => {
    const snapshot = createSnapshot();
    const sheet = snapshot.sheets[0];
    sheet.cells['0:0'] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'blank' },
      link: { target: 'https://example.com/' },
    };
    sheet.cells['1:0'] = {
      rowId: 1,
      columnId: 0,
      input: { type: 'blank' },
      note: 'note',
      link: { target: 'https://example.com/' },
    };
    const file = encodeWorkbookFile(snapshot);
    const { workbook, report: result } = toSheetJS(file);
    expect(result.issues).toContainEqual(
      expect.objectContaining({
        code: 'BLANK_LINK_INTEROP',
        action: 'approximated',
        sheetId: sheet.id,
        address: 'A1',
        count: 2,
        message: expect.stringContaining('some readers may expose the target as the cell value'),
      }),
    );
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
    const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    const reopened = XLSX.read(bytes, { type: 'buffer', sheetStubs: true });
    const imported = fromSheetJS(reopened).file;
    const reader = new FileReader(imported);
    try {
      expect(reader.getCell(imported.sheets[0].id, 0, 0)).toMatchObject({
        input: { type: 'blank' },
        link: { target: 'https://example.com/' },
      });
      expect(reader.getCell(imported.sheets[0].id, 1, 0)).toMatchObject({
        input: { type: 'blank' },
        note: 'note',
        link: { target: 'https://example.com/' },
      });
    } finally {
      reader.dispose();
    }
  });
  it('reports potential empty-text loss with location and blocks strict export', () => {
    const snapshot = createSnapshot();
    const sheet = snapshot.sheets[0];
    sheet.cells['0:0'] = { rowId: 0, columnId: 0, input: { type: 'string', value: '' } };
    sheet.cells['1:0'] = { rowId: 1, columnId: 0, input: { type: 'string', value: '' } };
    sheet.cells['2:0'] = { rowId: 2, columnId: 0, input: { type: 'blank' } };
    const file = encodeWorkbookFile(snapshot);
    const { workbook, report: result } = toSheetJS(file);
    expect(workbook.Sheets[sheet.name].A1).toMatchObject({ t: 's', v: '' });
    expect(result.issues).toContainEqual(
      expect.objectContaining({
        code: 'EMPTY_TEXT_INTEROP',
        action: 'approximated',
        sheetId: sheet.id,
        address: 'A1',
        count: 2,
      }),
    );
    expect(result.summary.approximated).toBe(2);
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
  it('rejects malformed historical reports instead of silently losing their warnings', () => {
    const valid = report();
    add(valid, issue({ action: 'dropped', count: 2 }));
    for (const prior of [
      { summary: valid.summary },
      { ...valid, summary: { ...valid.summary, dropped: -1 } },
      { ...valid, summary: { ...valid.summary, dropped: 1 } },
      { ...valid, issues: [issue({ action: 'preserved', count: 0 })] },
      { ...valid, issues: [{ ...issue(), action: 'unknown' }] },
      { ...valid, issues: [{ ...issue({ action: 'dropped' }), message: null }] },
      { ...valid, issues: [], omittedIssueOccurrences: 0 },
      { ...valid, omittedIssueOccurrences: -1 },
      { ...valid, omittedIssueOccurrences: 1 },
      { ...valid, issues: Array.from({ length: 201 }, () => issue()) },
    ]) {
      const file = createWorkbookFile();
      file.extensions['opensheet.importReport'] = prior;
      for (const unsupported of ['report', 'strict'] as const)
        expect(() => toSheetJS(file, { unsupported })).toThrow(
          'Invalid original import compatibility report',
        );
    }
  });
  it('rejects malformed historical reports before reading workbook pages', () => {
    const file = createWorkbookFile();
    file.extensions['opensheet.importReport'] = { issues: [], summary: { dropped: -1 } };
    Object.defineProperty(file.sheets[0], 'blocks', {
      get() {
        throw new Error('page data was accessed');
      },
    });
    for (const unsupported of ['report', 'strict'] as const) {
      expect(() => toSheetJS(file, { unsupported })).toThrow(
        'Invalid original import compatibility report',
      );
      expect(() => writeXlsx(file, XLSX, { unsupported })).toThrow(
        'Invalid original import compatibility report',
      );
    }
  });
  it('rejects invalid counts and actions from imported extension data', () => {
    for (const count of [-1, 0, 0.5, Infinity, Number.MAX_SAFE_INTEGER + 1])
      expect(() => add(report(), issue({ count }))).toThrow('Invalid compatibility issue');
    expect(() => add(report(), issue({ action: 'unknown' as Issue['action'] }))).toThrow(
      'Invalid compatibility issue',
    );
    const result = report();
    add(result, issue({ count: Number.MAX_SAFE_INTEGER }));
    expect(() => add(result, issue())).toThrow('exceeds safe range');
  });

  it('preserves aggregated and omitted import losses during subsequent exports', () => {
    const file = createWorkbookFile();
    const original = report();
    add(original, issue({ action: 'dropped', count: 7 }));
    for (let i = 1; i < 200; i++) add(original, issue({ code: `F${i}` }));
    add(original, issue({ action: 'blocked', count: 3 }));
    file.extensions['opensheet.importReport'] = original;
    const exported = toSheetJS(file).report;
    expect(exported.summary.dropped).toBe(7);
    expect(exported.summary.blocked).toBe(3);
    expect(exported.issues).toContainEqual(
      expect.objectContaining({ code: 'FEATURE', count: 3, action: 'blocked' }),
    );
    expect(exported.issues).toContainEqual(
      expect.objectContaining({ code: 'IMPORT_DETAILS_OMITTED', count: 7, action: 'dropped' }),
    );
    expect(() => toSheetJS(file, { unsupported: 'strict' })).toThrow();
  });

  it('retains totals for blockers omitted after the detail list contains only blockers', () => {
    const file = createWorkbookFile(),
      original = report();
    for (let i = 0; i < 201; i++) add(original, issue({ code: `BLOCK${i}`, action: 'blocked' }));
    expect(original.omittedIssueOccurrences).toBe(1);
    file.extensions['opensheet.importReport'] = original;
    const exported = toSheetJS(file).report;
    expect(exported.summary.blocked).toBe(201);
    expect(exported.omittedIssueOccurrences).toBe(1);
  });

  it.each(['approximated', 'dropped', 'blocked'] as const)(
    'rejects %s even after preserved details exhaust the cap',
    (action) => {
      const result = report();
      for (let i = 0; i < 200; i++) add(result, issue({ sheetId: `sheet-${i}` }));
      add(result, issue({ action }));
      expect(result.issues).toHaveLength(200);
      expect(result.summary.exact).toBe(200);
      expect(result.summary[action]).toBe(1);
      expect(result.omittedIssueOccurrences).toBe(1);
      expect(() => finish(result, { unsupported: 'strict' })).toThrow(
        'Strict compatibility check failed',
      );
      expect(() => finish(result, { unsupported: 'report' })).not.toThrow();
    },
  );

  it('does not merge different actions or severities under the same code', () => {
    const result = report();
    add(result, issue());
    add(result, issue({ action: 'dropped', severity: 'warning' }));
    add(result, issue({ action: 'dropped', severity: 'error' }));
    expect(result.issues.map((entry) => [entry.action, entry.severity])).toEqual([
      ['preserved', 'info'],
      ['dropped', 'warning'],
      ['dropped', 'error'],
    ]);
    expect(() => finish(result, { unsupported: 'strict' })).toThrow();
  });

  it('counts retained groups after saturation and preserves the first address', () => {
    const result = report();
    const input = issue({ address: 'A1' });
    add(result, input);
    for (let i = 1; i < 200; i++) add(result, issue({ code: `F${i}` }));
    add(result, issue({ address: 'B1' }));
    add(result, issue({ code: 'OMITTED' }));
    add(result, issue({ code: 'OMITTED' }));
    expect(result.issues[0]).toMatchObject({ count: 2, address: 'A1' });
    expect(input.count).toBeUndefined();
    expect(result.omittedIssueOccurrences).toBe(2);
    expect(result.summary.exact).toBe(203);
    expect(() => finish(result, { unsupported: 'strict' })).not.toThrow();
  });

  it('retains a new blocker when earlier warning groups fill the detail cap', () => {
    const result = report();
    add(result, issue({ code: 'FIRST', action: 'approximated', count: 3 }));
    for (let i = 1; i < 200; i++) add(result, issue({ code: `F${i}`, action: 'dropped' }));
    add(result, issue({ code: 'BLOCKER', action: 'blocked', address: 'B7' }));
    add(result, issue({ code: 'BLOCKER', action: 'blocked', address: 'C8' }));
    expect(result.issues).toHaveLength(200);
    expect(result.issues).toContainEqual(
      expect.objectContaining({ code: 'BLOCKER', address: 'B7', count: 2 }),
    );
    expect(result.issues.some((entry) => entry.code === 'FIRST')).toBe(false);
    expect(result.omittedIssueOccurrences).toBe(3);
    expect(result.summary).toEqual({ exact: 0, approximated: 3, dropped: 199, blocked: 2 });
    expect(() => finish(result, { unsupported: 'strict' })).toThrow(
      'Strict compatibility check failed',
    );
  });
});
