import { describe, it, expect, vi } from 'vitest';
import fc from 'fast-check';
import {
  createWorkbook,
  Workbook,
  createSnapshot,
  validateSnapshot,
  address,
  parseRange,
  formatValue,
  key,
  type CommandEnvelope,
} from '@opensheetjs/core';
import { normalizedSnapshot } from '@opensheetjs/testing';
describe('workbook contracts', () => {
  it('preserves blank, empty string, zero, false, and literal formula text', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('A1:E1').setValues([[null, '', 0, false, '=1+1']]);
    expect(s.range('A1:E1').getValues()).toEqual([[null, '', 0, false, '=1+1']]);
    expect(new Workbook(b.toJSON()).getSheets()[0].range('A1:E1').getValues()).toEqual([
      [null, '', 0, false, '=1+1'],
    ]);
  });
  it('rejects ragged, nonfinite and oversized writes atomically', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0],
      before = b.toJSON();
    expect(() => s.range('A1:B2').setValues([[1], [2, 3]])).toThrow();
    expect(() => s.range('A1:B1').setValues([[1, NaN]])).toThrow();
    expect(() => s.range('A1:XFD1048576')).toThrow();
    expect(b.toJSON()).toEqual(before);
  });
  it('groups commands and rolls back entire failed transactions', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0],
      fn = vi.fn();
    b.onCommit(fn);
    b.transaction({}, () => {
      s.range('A1').setValues([[1]]);
      s.range('A2').setValues([[2]]);
    });
    expect(fn).toHaveBeenCalledTimes(1);
    expect(b.revision).toBe(1);
    b.undo();
    expect(s.range('A1:A2').getValues()).toEqual([[null], [null]]);
    b.redo();
    expect(s.range('A1:A2').getValues()).toEqual([[1], [2]]);
    const before = b.toJSON();
    expect(() =>
      b.transaction({}, () => {
        s.range('A1').setValues([[99]]);
        s.deleteRows(0, 100);
      }),
    ).toThrow();
    expect(b.toJSON()).toEqual(before);
  });
  it('poisons a transaction even when the caller catches an invalid command', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    expect(() =>
      b.transaction({}, () => {
        s.range('A1').setValues([[7]]);
        try {
          s.deleteRows(0, 100);
        } catch {}
      }),
    ).toThrow();
    expect(s.range('A1').getValues()).toEqual([[null]]);
  });
  it('rejects async and nested transactions', () => {
    const b = createWorkbook();
    expect(() => b.transaction({}, async () => {})).toThrow();
    expect(() => b.transaction({}, () => b.transaction({}, () => {}))).toThrow();
  });
  it('does not expose mutable model data', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    expect(() => b.sheetData(s.id).rowOrder.push('evil')).toThrow();
    const snap = b.toJSON();
    snap.sheets[0].name = 'Changed';
    expect(s.name).toBe('Sheet1');
  });
  it('supports revision preconditions and retry deduplication', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    const cmd: CommandEnvelope = {
      protocolVersion: 1,
      commandId: 'retry',
      workbookId: b.id,
      baseRevision: 0,
      type: 'core.cells.set',
      payload: {
        sheetId: s.id,
        cells: [{ row: 0, column: 0, input: { type: 'number', value: 5 } }],
      },
    };
    const a = b.execute(cmd);
    expect(b.execute(cmd)).toEqual(a);
    expect(b.revision).toBe(1);
    expect(() => b.execute({ ...cmd, commandId: 'stale' })).toThrow(/changed/);
    expect(() => b.execute({ ...cmd, payload: { sheetId: s.id, cells: [] } })).toThrow(/reused/);
  });
  it('enforces read-only even through commands and undo', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('A1').setValues([[1]]);
    b.setReadOnly(true);
    expect(() => s.range('A1').setValues([[2]])).toThrow(/read only/);
    expect(() => b.undo()).toThrow();
    b.setReadOnly(false);
    expect(b.undo()).toBe(true);
  });
  it('retains stable row identities and rewrites cross-sheet references', () => {
    const b = createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Summary' }] }),
      [s, t] = b.getSheets();
    s.range('A1').setValues([[8]]);
    t.range('A1').setFormulas([["'Data'!$A$1*2"]]);
    const id = b.getCell(s.id, 0, 0)!.rowId;
    s.insertRows(0);
    expect(b.getCell(s.id, 1, 0)!.rowId).toBe(id);
    expect(t.range('A1').getValues()).toEqual([[16]]);
    expect(t.range('A1').getFormulas()[0][0]).toContain('$A$2');
    b.renameSheet(s.id, 'New data');
    expect(t.range('A1').getValues()).toEqual([[16]]);
    s.deleteRows(1);
    expect(t.range('A1').getValues()).toEqual([[{ error: '#REF!' }]]);
  });
  it('rejects unsafe structural edits with unsupported references', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('A1').setFormulas([['ExternalName+1']]);
    const before = b.toJSON();
    expect(() => s.insertRows(1)).toThrow(/safely/);
    expect(b.toJSON()).toEqual(before);
  });
  it('preserves strings during formula reference rewrites', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('B1').setFormulas([['"A1"&A1']]);
    s.range('A1').setValues([['x']]);
    s.insertRows(0);
    expect(s.range('B2').getValues()).toEqual([['A1x']]);
  });
  it('merge refuses data loss and covered-cell writes, and can be undone', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('A1:B1').setValues([[1, 2]]);
    expect(() => s.range('A1:B1').merge()).toThrow();
    s.range('A1:B1').merge({ discardCoveredValues: true });
    expect(() => s.range('B1').setValues([[3]])).toThrow(/covered/);
    expect(() => s.range('A1').clear()).toThrow(/cuts/);
    b.undo();
    expect(s.range('A1:B1').getValues()).toEqual([[1, 2]]);
  });
  it('clear preserves style, clear-all removes it', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('A1').setValues([[1]]);
    s.range('A1').setStyle({ bold: true });
    s.range('A1').clear();
    expect(b.getCell(s.id, 0, 0)?.styleId).toBeTruthy();
    s.range('A1').clear({ all: true });
    expect(b.getCell(s.id, 0, 0)).toBeUndefined();
  });
  it('sort is stable, moves formulas and does not move identities', () => {
    const b = createWorkbook(),
      s = b.getSheets()[0];
    s.range('A1:B3').setValues([
      [3, 30],
      [1, 10],
      [2, 20],
    ]);
    s.range('C1:C3').setFormulas([['B1*2'], ['B2*2'], ['B3*2']]);
    const rows = [...b.sheetData(s.id).rowOrder];
    s.sort('A1:C3', 0);
    expect(s.range('A1:C3').getValues()).toEqual([
      [1, 10, 20],
      [2, 20, 40],
      [3, 30, 60],
    ]);
    expect(b.sheetData(s.id).rowOrder).toEqual(rows);
  });
  it('validates malicious JSON, duplicate ids and out-of-range cells', () => {
    const snapshot = createSnapshot();
    const unsafe = JSON.parse(
      JSON.stringify(snapshot).replace(
        '"extensions":{}',
        '"extensions":{"__proto__":{"bad":true}}',
      ),
    );
    expect(() => validateSnapshot(unsafe)).toThrow();
    const copy = structuredClone(snapshot);
    copy.sheets[0].rowOrder[1] = copy.sheets[0].rowOrder[0];
    expect(() => validateSnapshot(copy)).toThrow(/Duplicate/);
    expect(() => validateSnapshot({ ...snapshot, schemaVersion: 2 })).toThrow();
  });
  it('invalidates handles on deletion and disposal', () => {
    const b = createWorkbook(),
      s = b.addSheet('Temporary');
    b.removeSheet(s.id);
    expect(() => s.range('A1')).toThrow();
    b.dispose();
    b.dispose();
    expect(() => b.toJSON()).toThrow(/disposed/);
  });
  it('reports the date-system leap day correctly', () => {
    expect(formatValue(59, 'yyyy-mm-dd')).toBe('1900-02-28');
    expect(formatValue(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(formatValue(61, 'yyyy-mm-dd')).toBe('1900-03-01');
    expect(formatValue(0, 'yyyy-mm-dd', '1904')).toBe('1904-01-01');
  });
  it('round-trips writes through undo across arbitrary finite values', () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.oneof(fc.integer(), fc.string({ maxLength: 50 }), fc.boolean(), fc.constant(null)),
          { minLength: 1, maxLength: 40 },
        ),
        (values) => {
          const b = createWorkbook(),
            s = b.getSheets()[0],
            before = normalizedSnapshot(b.toJSON());
          s.range(`A1:A${values.length}`).setValues(values.map((v) => [v]));
          b.undo();
          expect(normalizedSnapshot(b.toJSON())).toEqual(before);
          b.redo();
          expect(s.range(`A1:A${values.length}`).getValues()).toEqual(values.map((v) => [v]));
        },
      ),
      { numRuns: 50 },
    );
  });
  it('round-trips addresses', () => {
    fc.assert(
      fc.property(fc.integer({ min: 0, max: 99999 }), fc.integer({ min: 0, max: 999 }), (r, c) => {
        expect(parseRange(address(r, c))).toEqual({
          startRow: r,
          endRow: r + 1,
          startColumn: c,
          endColumn: c + 1,
        });
      }),
    );
  });
});
