import { expect, it } from 'vitest';
import { Engine } from '../packages/core/src/engine.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { Database } from '../packages/core/src/storage.js';

it.each([[[0, '0']], [[false, 'false']], [[null, 'null']], [['a'.repeat(101), 'second']]])(
  'rejects invalid persisted style identities %j',
  async (ids) => {
    const db = await Database.open(`style-identity-${crypto.randomUUID()}`);
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 3, columns: 3 }] });
    snapshot.styles = { first: { bold: true }, second: { italic: true } };
    snapshot.sheets[0].cells['0:0'] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'number', value: 1 },
      styleId: 'first',
    };
    snapshot.sheets[0].cells['1:0'] = {
      rowId: 1,
      columnId: 0,
      input: { type: 'number', value: 2 },
      styleId: 'second',
    };
    let engine = new Engine(db, snapshot.workbookId);
    try {
      await engine.open(snapshot);
      const head = await db.get<{ meta: string }>('heads', snapshot.workbookId);
      const metadata = await db.get<{ styleIds: unknown[]; snapshot: { styles: unknown[] } }>(
        'records',
        head!.meta,
      );
      expect(metadata!.styleIds).toHaveLength(2);
      metadata!.styleIds = ids;
      await db.write('records', [[head!.meta, metadata]]);
      engine.close();
      const reopened = await Database.open(db.db.name);
      engine = new Engine(reopened, snapshot.workbookId);
      await expect(engine.open()).rejects.toMatchObject({
        code: 'CORRUPT_STORAGE',
        message: 'Invalid stored style identities',
      });
    } finally {
      engine.close();
    }
  },
);

it('reopens the maximum-length valid style identity', async () => {
  const db = await Database.open(`style-identity-valid-${crypto.randomUUID()}`);
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 3, columns: 3 }] });
  const id = 'a'.repeat(100);
  snapshot.styles = { [id]: { bold: true } };
  snapshot.sheets[0].cells['0:0'] = {
    rowId: 0,
    columnId: 0,
    input: { type: 'number', value: 1 },
    styleId: id,
  };
  let engine = new Engine(db, snapshot.workbookId);
  try {
    await engine.open(snapshot);
    engine.close();
    engine = new Engine(await Database.open(db.db.name), snapshot.workbookId);
    const view = await engine.open();
    expect(view.snapshot.styles[id]).toEqual({ bold: true });
  } finally {
    engine.close();
  }
});
