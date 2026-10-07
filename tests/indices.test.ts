import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { AxisIndex } from '../packages/core/src/axis-index.js';
import { ReferenceIndex } from '../packages/core/src/reference-index.js';
import { createWorkbook } from '@opensheetjs/core';
import { Engine } from '../packages/core/src/engine.js';
import { Database } from '../packages/core/src/storage.js';
import { createSnapshot } from '../packages/core/src/model.js';
describe('incremental indices and persistence', () => {
  it('matches an independent array model through randomized chunk splits and deletions', () => {
    fc.assert(
      fc.property(
        fc.array(
          fc.record({
            at: fc.nat(10000),
            count: fc.integer({ min: 1, max: 8 }),
            insert: fc.boolean(),
          }),
          { maxLength: 100 },
        ),
        (operations) => {
          let ids = Array.from({ length: 300 }, (_, i) => i),
            index = new AxisIndex(ids),
            next = 300;
          for (const operation of operations) {
            const old = index.clone(),
              before = [...ids],
              at = operation.at % (ids.length + 1);
            if (operation.insert) {
              const inserted = Array.from({ length: operation.count }, () => next++);
              index.splice(at, 0, inserted);
              ids.splice(at, 0, ...inserted);
            } else {
              const count = Math.min(operation.count, ids.length - at);
              index.splice(at, count);
              ids.splice(at, count);
            }
            expect(index.toArray()).toEqual(ids);
            expect(old.toArray()).toEqual(before);
            ids.forEach((id, position) => {
              expect(index.get(id)).toBe(position);
              expect(index.at(position)).toBe(id);
            });
            for (const id of before.filter((id) => !ids.includes(id)))
              expect(index.get(id)).toBeUndefined();
          }
        },
      ),
      { numRuns: 30 },
    );
  });
  it('matches naive range dependency intersection, including removal', () => {
    const index = new ReferenceIndex(),
      refs = Array.from({ length: 300 }, (_, i) => ({
        startRow: i % 50,
        endRow: (i % 50) + 1 + (i % 7),
        startColumn: i % 13,
        endColumn: (i % 13) + 1 + (i % 5),
      }));
    refs.forEach((range, i) => index.add(String(i), [{ sheet: 's', range }]));
    for (let i = 0; i < 300; i += 3) index.delete(String(i));
    for (let row = 0; row < 55; row++) {
      const query = { startRow: row, endRow: row + 2, startColumn: 3, endColumn: 8 };
      const expected = refs
        .flatMap((r, i) =>
          i % 3 !== 0 &&
          r.startRow < query.endRow &&
          r.endRow > query.startRow &&
          r.startColumn < query.endColumn &&
          r.endColumn > query.startColumn
            ? [String(i)]
            : [],
        )
        .sort();
      expect([...index.query('s', query)].sort()).toEqual(expected);
    }
  });
  it('invalidates cached formulas transitively, including constants and range dependencies', async () => {
    const book = await createWorkbook();
    try {
      const sh = book.getSheets()[0];
      await sh.range('A1').setValues([[2]]);
      await sh.range('B1:D1').setFormulas([['A1*2', 'SUM(A1:B1)', '1+1']]);
      expect(await sh.range('B1:D1').getValues()).toEqual([[4, 6, 2]]);
      await sh.range('A1').setValues([[3]]);
      expect(await sh.range('B1:C1').getValues()).toEqual([[6, 9]]);
      await sh.range('D1').setFormulas([['3+3']]);
      expect(await sh.range('D1').getValues()).toEqual([[6]]);
    } finally {
      book.dispose();
    }
  });
  it('compresses persisted pages without changing revision and retains undo/redo after reopen', async () => {
    const dbName = `compression-${crypto.randomUUID()}`,
      snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 100, columns: 32 }] });
    const engine = new Engine(await Database.open(dbName), snapshot.workbookId);
    await engine.open(snapshot);
    const sheet = snapshot.sheets[0];
    await engine.execute([
      {
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: Array.from({ length: 64 * 32 }, (_, i) => ({
            row: Math.floor(i / 32),
            column: i % 32,
            input: { type: 'string', value: 'repeated value '.repeat(8) },
          })),
        },
      },
    ]);
    const before = engine.view(),
      revision = before.revision;
    await engine.compact(10000);
    expect(engine.view().revision).toBe(revision);
    expect(engine.view().bytes).toBeLessThan(before.bytes);
    engine.close();
    const reopened = new Engine(await Database.open(dbName), snapshot.workbookId);
    await reopened.open();
    expect((await reopened.cell(sheet.id, 0, 0))?.input).toEqual({
      type: 'string',
      value: 'repeated value '.repeat(8),
    });
    await reopened.history(false);
    expect(await reopened.cell(sheet.id, 0, 0)).toBeUndefined();
    await reopened.history(true);
    expect((await reopened.cell(sheet.id, 63, 31))?.input.type).toBe('string');
    reopened.close();
  });
  it('offers field reads and occupied-cell cursors that reject changed revisions', async () => {
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 1000, columns: 100 }] });
    try {
      const sh = book.getSheets()[0];
      await sh.range('A1:B1').setValues([[0, false]]);
      await sh.range('A1000').setValues([['far']]);
      expect(
        (
          await book.readCells(
            sh.id,
            { startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 },
            { fields: ['input'] },
          )
        ).cells,
      ).toEqual([
        { row: 0, column: 0, input: { type: 'number', value: 0 } },
        { row: 0, column: 1, input: { type: 'boolean', value: false } },
      ]);
      const first = await book.scanCells(sh.id, { limit: 1 });
      expect(first.cursor).toBeDefined();
      const second = await book.scanCells(sh.id, { cursor: first.cursor, limit: 1 });
      expect(second.cells[0].column).toBe(1);
      await sh.range('C1').setValues([[7]]);
      await expect(book.scanCells(sh.id, { cursor: second.cursor })).rejects.toMatchObject({
        code: 'REVISION_CONFLICT',
      });
    } finally {
      book.dispose();
    }
  });
  it('reuses unchanged metadata through scalar undo/redo without losing invalidation', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 100000, columns: 100 }] });
    const engine = new Engine(
      await Database.open(`history-${crypto.randomUUID()}`),
      snapshot.workbookId,
    );
    try {
      await engine.open(snapshot);
      const rowOrder = engine.view().snapshot.sheets[0].rowOrder;
      const sheetId = snapshot.sheets[0].id;
      await engine.execute([
        {
          type: 'core.cells.set',
          payload: {
            sheetId,
            cells: [
              { row: 0, column: 0, input: { type: 'number', value: 7 } },
              { row: 0, column: 1, input: { type: 'formula', expression: 'A1*2' } },
            ],
          },
        },
      ]);
      await engine.history(false);
      expect(engine.view().snapshot.sheets[0].rowOrder).toBe(rowOrder);
      expect(await engine.cell(sheetId, 0, 0)).toBeUndefined();
      await engine.history(true);
      expect(engine.view().snapshot.sheets[0].rowOrder).toBe(rowOrder);
      expect(
        (await engine.read(sheetId, { startRow: 0, endRow: 1, startColumn: 1, endColumn: 2 }))
          .calculated['0:1'],
      ).toBe(14);
    } finally {
      engine.close();
    }
  });
  it('does not read or rewrite unrelated numeric pages during insertion', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 1000, columns: 100 }] });
    const db = await Database.open(`insert-${crypto.randomUUID()}`),
      engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    await engine.generate(10000, 256 * 1024 * 1024);
    const before = await db.get<{ directory: Record<string, string> }>(
      'heads',
      snapshot.workbookId,
    );
    await engine.execute([
      {
        type: 'core.axis.insert',
        payload: { sheetId: snapshot.sheets[0].id, axis: 'row', index: 0, ids: [100] },
      },
    ]);
    const after = await db.get<{ directory: Record<string, string> }>('heads', snapshot.workbookId);
    expect(after!.directory).toEqual(before!.directory);
    expect((await engine.cell(snapshot.sheets[0].id, 1, 0))?.input).toEqual({
      type: 'number',
      value: 0,
    });
    engine.close();
  });
});
