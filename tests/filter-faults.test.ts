import { describe, expect, it, vi } from 'vitest';
import { createSnapshot } from '../packages/core/src/model.js';
import { Engine } from '../packages/core/src/engine.js';
import { Database } from '../packages/core/src/storage.js';
import { key } from '@opensheetjs/core';

describe('filter failure isolation', () => {
  it('keeps the committed view readable while a staged filter waits', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 3, columns: 1 }] });
    const sheet = snapshot.sheets[0];
    sheet.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } };
    sheet.filter = {
      mode: 'all',
      headerRows: 0,
      criteria: [{ column: 0, operator: 'gt', value: 0 }],
    };
    const db = await Database.open(`filter-isolation-${crypto.randomUUID()}`);
    const engine = new Engine(db, snapshot.workbookId, 0);
    await engine.open(snapshot);
    const before = engine.view();
    let release!: () => void;
    let entered!: () => void;
    const blocked = new Promise<void>((resolve) => {
      release = resolve;
    });
    const started = new Promise<void>((resolve) => {
      entered = resolve;
    });
    const stagedCell = (engine as any).stagedCell.bind(engine);
    let first = true;
    const spy = vi.spyOn(engine as any, 'stagedCell').mockImplementation(async (...args: any[]) => {
      const result = await stagedCell(...args);
      if (first) {
        first = false;
        entered();
        await blocked;
      }
      return result;
    });
    try {
      const write = engine.execute([
        {
          type: 'core.cells.set',
          payload: {
            sheetId: sheet.id,
            cells: [{ row: 0, column: 0, input: { type: 'number', value: 9 } }],
          },
        },
      ]);
      await started;
      expect(engine.view().revision).toBe(before.revision);
      expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({ type: 'number', value: 1 });
      release();
      await write;
      expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({ type: 'number', value: 9 });
    } finally {
      release();
      spy.mockRestore();
      engine.close();
    }
  });

  it.each(['cancel', 'quota'] as const)(
    'retains data, visibility and history after %s during a staged filter edit',
    async (fault) => {
      const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 100, columns: 2 }] });
      const sheet = snapshot.sheets[0];
      for (let row = 0; row < 100; row++) {
        sheet.cells[key(row, 0)] = {
          rowId: row,
          columnId: 0,
          input: { type: 'number', value: row },
        };
        sheet.cells[key(row, 1)] = {
          rowId: row,
          columnId: 1,
          input: { type: 'formula', expression: `A${row + 1}*2` },
        };
      }
      sheet.filter = {
        mode: 'all',
        headerRows: 0,
        criteria: [{ column: 1, operator: 'gt', value: 100 }],
      };
      const name = `filter-fault-${crypto.randomUUID()}`;
      let db = await Database.open(name),
        engine = new Engine(db, snapshot.workbookId, 0);
      await engine.open(snapshot);
      const before = engine.view();
      const cell = (engine as any).stagedCell.bind(engine),
        write = db.write.bind(db);
      let reads = 0;
      const reader = vi
        .spyOn(engine as any, 'stagedCell')
        .mockImplementation(async (...args: any[]) => {
          const result = await cell(...args);
          if (fault === 'cancel' && ++reads === 10) engine.cancel();
          return result;
        });
      const writer = vi.spyOn(db, 'write').mockImplementation(async (...args) => {
        if (fault === 'quota' && args[0] === 'heads')
          throw new DOMException('Injected quota failure', 'QuotaExceededError');
        return write(...args);
      });
      try {
        await expect(
          engine.execute([
            {
              type: 'core.cells.set',
              payload: {
                sheetId: sheet.id,
                cells: [{ row: 0, column: 0, input: { type: 'number', value: 999 } }],
              },
            },
            {
              type: 'core.sheet.filter',
              payload: {
                sheetId: sheet.id,
                filter: {
                  mode: 'any',
                  headerRows: 0,
                  criteria: [{ column: 1, operator: 'lt', value: 100 }],
                },
              },
            },
          ]),
        ).rejects.toMatchObject(
          fault === 'cancel' ? { code: 'ABORTED' } : { name: 'QuotaExceededError' },
        );
        reader.mockRestore();
        writer.mockRestore();
        engine.begin();
        expect(engine.view().snapshot).toEqual(before.snapshot);
        expect(engine.view().bytes).toBe(before.bytes);
        expect(engine.view().filtered).toEqual(before.filtered);
        expect(engine.view().canUndo).toBe(before.canUndo);
        expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({ type: 'number', value: 0 });
        engine.close();
        db = await Database.open(name);
        engine = new Engine(db, snapshot.workbookId, 0);
        await engine.open();
        expect(engine.view().snapshot).toEqual(before.snapshot);
        expect(engine.view().bytes).toBe(before.bytes);
        expect(engine.view().filtered).toEqual(before.filtered);
        expect(engine.view().canUndo).toBe(before.canUndo);
        expect((await engine.cell(sheet.id, 0, 0))?.input).toEqual({ type: 'number', value: 0 });
      } finally {
        reader.mockRestore();
        writer.mockRestore();
        engine.close();
      }
    },
  );

  it('cancels history filter recalculation without changing the live mask or history position', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 100, columns: 1 }] });
    const sheet = snapshot.sheets[0];
    sheet.filter = { mode: 'all', headerRows: 0, criteria: [{ column: 0, operator: 'isBlank' }] };
    const db = await Database.open(`filter-history-${crypto.randomUUID()}`);
    const engine = new Engine(db, snapshot.workbookId);
    await engine.open(snapshot);
    await engine.execute([
      {
        type: 'core.cells.set',
        payload: {
          sheetId: sheet.id,
          cells: [{ row: 0, column: 0, input: { type: 'number', value: 1 } }],
        },
      },
    ]);
    const before = engine.view(),
      cell = (engine as any).stagedCell.bind(engine);
    let reads = 0;
    const spy = vi.spyOn(engine as any, 'stagedCell').mockImplementation(async (...args: any[]) => {
      const result = await cell(...args);
      if (++reads === 10) engine.cancel();
      return result;
    });
    try {
      await expect(engine.history(false)).rejects.toMatchObject({ code: 'ABORTED' });
      expect(engine.view().snapshot.revision).toBe(before.snapshot.revision);
      expect(engine.view().filtered).toEqual(before.filtered);
      expect(engine.view().canUndo).toBe(true);
      expect(engine.view().canRedo).toBe(false);
      spy.mockRestore();
      await engine.history(false);
      expect(engine.view().filtered[sheet.id]).toEqual([]);
      await engine.history(true);
      expect(engine.view().filtered[sheet.id]).toEqual([0]);
    } finally {
      spy.mockRestore();
      engine.close();
    }
  });
});
