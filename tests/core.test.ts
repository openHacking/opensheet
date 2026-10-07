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
  it('preserves blank, empty string, zero, false, and literal formula text', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1:E1').setValues([[null, '', 0, false, '=1+1']]);
    expect(await s.range('A1:E1').getValues()).toEqual([[null, '', 0, false, '=1+1']]);
    expect(
      await (await new Workbook(await b.toJSON()).ready())
        .getSheets()[0]
        .range('A1:E1')
        .getValues(),
    ).toEqual([[null, '', 0, false, '=1+1']]);
  });
  it('rejects ragged, nonfinite and oversized writes atomically', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0],
      before = await b.toJSON();
    await expect(async () => await s.range('A1:B2').setValues([[1], [2, 3]])).rejects.toThrow();
    await expect(async () => await s.range('A1:B1').setValues([[1, NaN]])).rejects.toThrow();
    expect(() => s.range('A1:XFD1048576')).toThrow();
    expect(await b.toJSON()).toEqual(before);
  });
  it('groups commands and rolls back entire failed transactions', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0],
      fn = vi.fn();
    b.onCommit(fn);
    await b.transaction({}, async (b) => {
      await b
        .getSheetById(s.id)!
        .range('A1')
        .setValues([[1]]);
      await b
        .getSheetById(s.id)!
        .range('A2')
        .setValues([[2]]);
    });
    expect(fn).toHaveBeenCalledTimes(1);
    expect(b.revision).toBe(1);
    await b.undo();
    expect(await s.range('A1:A2').getValues()).toEqual([[null], [null]]);
    await b.redo();
    expect(await s.range('A1:A2').getValues()).toEqual([[1], [2]]);
    const before = await b.toJSON();
    await expect(
      async () =>
        await b.transaction({}, async (b) => {
          await b
            .getSheetById(s.id)!
            .range('A1')
            .setValues([[99]]);
          await b.getSheetById(s.id)!.deleteRows(0, 100);
        }),
    ).rejects.toThrow();
    expect(await b.toJSON()).toEqual(before);
  });
  it('poisons a transaction even when the caller catches an invalid command', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await expect(
      async () =>
        await b.transaction({}, async (b) => {
          await b
            .getSheetById(s.id)!
            .range('A1')
            .setValues([[7]]);
          try {
            await b.getSheetById(s.id)!.deleteRows(0, 100);
          } catch {}
        }),
    ).rejects.toThrow();
    expect(await s.range('A1').getValues()).toEqual([[null]]);
  });
  it('supports async callbacks and rejects nested transactions', async () => {
    const b = await createWorkbook();
    await b.transaction({}, async (b) => {
      await b
        .getSheets()[0]
        .range('A1')
        .setValues([[5]]);
    });
    expect(await b.getSheets()[0].range('A1').getValues()).toEqual([[5]]);
    await expect(
      async () => await b.transaction({}, async (b) => await b.transaction({}, () => {})),
    ).rejects.toThrow();
  });
  it('does not expose mutable model data', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    expect(() => b.sheetData(s.id).rowOrder.push('evil')).toThrow();
    const snap = await b.toJSON();
    snap.sheets[0].name = 'Changed';
    expect(s.name).toBe('Sheet1');
  });
  it('supports revision preconditions and retry deduplication', async () => {
    const b = await createWorkbook(),
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
    const a = await b.execute(cmd);
    expect(await b.execute(cmd)).toEqual(a);
    expect(b.revision).toBe(1);
    await expect(async () => await b.execute({ ...cmd, commandId: 'stale' })).rejects.toThrow(
      /changed/,
    );
    await expect(
      async () => await b.execute({ ...cmd, payload: { sheetId: s.id, cells: [] } }),
    ).rejects.toThrow(/reused/);
  });
  it('enforces read-only even through commands and undo', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1').setValues([[1]]);
    b.setReadOnly(true);
    await expect(async () => await s.range('A1').setValues([[2]])).rejects.toThrow(/read only/);
    await expect(async () => await b.undo()).rejects.toThrow();
    b.setReadOnly(false);
    expect(await b.undo()).toBe(true);
  });
  it('retains stable row identities and rewrites cross-sheet references', async () => {
    const b = await createWorkbook({ sheets: [{ name: 'Data' }, { name: 'Summary' }] }),
      [s, t] = b.getSheets();
    await s.range('A1').setValues([[8]]);
    await t.range('A1').setFormulas([["'Data'!$A$1*2"]]);
    const id = (await b.getCell(s.id, 0, 0))!.rowId;
    await s.insertRows(0);
    expect((await b.getCell(s.id, 1, 0))!.rowId).toBe(id);
    expect(await t.range('A1').getValues()).toEqual([[16]]);
    expect((await t.range('A1').getFormulas())[0][0]).toContain('$A$2');
    await b.renameSheet(s.id, 'New data');
    expect(await t.range('A1').getValues()).toEqual([[16]]);
    await s.deleteRows(1);
    expect(await t.range('A1').getValues()).toEqual([[{ error: '#REF!' }]]);
  });
  it('rejects unsafe structural edits with unsupported references', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1').setFormulas([['ExternalName+1']]);
    const before = await b.toJSON();
    await expect(async () => await s.insertRows(1)).rejects.toThrow(/safely/);
    expect(await b.toJSON()).toEqual(before);
  });
  it('preserves strings during formula reference rewrites', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('B1').setFormulas([['"A1"&A1']]);
    await s.range('A1').setValues([['x']]);
    await s.insertRows(0);
    expect(await s.range('B2').getValues()).toEqual([['A1x']]);
  });
  it('merge refuses data loss and covered-cell writes, and can be undone', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1:B1').setValues([[1, 2]]);
    await expect(async () => await s.range('A1:B1').merge()).rejects.toThrow();
    await s.range('A1:B1').merge({ discardCoveredValues: true });
    await expect(async () => await s.range('B1').setValues([[3]])).rejects.toThrow(/covered/);
    await expect(async () => await s.range('A1').clear()).rejects.toThrow(/cuts/);
    await b.undo();
    expect(await s.range('A1:B1').getValues()).toEqual([[1, 2]]);
  });
  it('clear preserves style, clear-all removes it', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1').setValues([[1]]);
    await s.range('A1').setStyle({ bold: true });
    await s.range('A1').clear();
    expect((await b.getCell(s.id, 0, 0))?.styleId).toBeTruthy();
    await s.range('A1').clear({ all: true });
    expect(await b.getCell(s.id, 0, 0)).toBeUndefined();
  });
  it('sort is stable, moves formulas and does not move identities', async () => {
    const b = await createWorkbook(),
      s = b.getSheets()[0];
    await s.range('A1:B3').setValues([
      [3, 30],
      [1, 10],
      [2, 20],
    ]);
    await s.range('C1:C3').setFormulas([['B1*2'], ['B2*2'], ['B3*2']]);
    const rows = [...b.sheetData(s.id).rowOrder];
    await s.sort('A1:C3', 0);
    expect(await s.range('A1:C3').getValues()).toEqual([
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
  it('invalidates handles on deletion and disposal', async () => {
    const b = await createWorkbook(),
      s = await b.addSheet('Temporary');
    await b.removeSheet(s.id);
    expect(() => s.range('A1')).toThrow();
    b.dispose();
    b.dispose();
    await expect(async () => await b.toJSON()).rejects.toThrow(/disposed/);
  });
  it('reports the date-system leap day correctly', () => {
    expect(formatValue(59, 'yyyy-mm-dd')).toBe('1900-02-28');
    expect(formatValue(60, 'yyyy-mm-dd')).toBe('1900-02-29');
    expect(formatValue(61, 'yyyy-mm-dd')).toBe('1900-03-01');
    expect(formatValue(0, 'yyyy-mm-dd', '1904')).toBe('1904-01-01');
  });
  it('round-trips writes through undo across arbitrary finite values', async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(
          fc.oneof(fc.integer(), fc.string({ maxLength: 50 }), fc.boolean(), fc.constant(null)),
          { minLength: 1, maxLength: 40 },
        ),
        async (values) => {
          const b = await createWorkbook(),
            s = b.getSheets()[0],
            before = normalizedSnapshot(await b.toJSON());
          await s.range(`A1:A${values.length}`).setValues(values.map((v) => [v]));
          await b.undo();
          expect(normalizedSnapshot(await b.toJSON())).toEqual(before);
          await b.redo();
          expect(await s.range(`A1:A${values.length}`).getValues()).toEqual(values.map((v) => [v]));
        },
      ),
      { numRuns: 50 },
    );
  });
  it('round-trips addresses', async () => {
    await fc.assert(
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
