import { expect, it, vi } from 'vitest';
import { createWorkbook, openWorkbook } from '@opensheetjs/core';
import { Database, completed } from '../packages/core/src/storage.js';

it('bounds cursor batches by bytes without skipping oversized or deleted records', async () => {
  const db = await Database.open(`scan-${crypto.randomUUID()}`);
  try {
    const entries: [string, Uint8Array][] = [100, 100, 900, 100, 100, 100].map((size, i) => [
      String(i),
      new Uint8Array(size),
    ]);
    await db.write('scratch', entries);
    const batches = [];
    for await (const batch of db.scan<Uint8Array>('scratch', IDBKeyRange.bound('0', '9'), 3, 250)) {
      batches.push(batch.map(([id]) => id));
      await db.write(
        'scratch',
        [],
        batch.map(([id]) => id),
      );
    }
    expect(batches).toEqual([['0', '1'], ['2'], ['3', '4'], ['5']]);
    const tx = db.db.transaction('scratch', 'readwrite');
    const done = completed(tx);
    tx.objectStore('scratch').put(new Uint8Array([9]), 'bad');
    await done;
    const read = async () => {
      for await (const _ of db.scan('scratch', IDBKeyRange.only('bad'), 3, 250)) {
        /* consume */
      }
    };
    await expect(read()).rejects.toMatchObject({ code: 'CORRUPT_STORAGE' });
  } finally {
    db.close();
  }
});

it.each(['quota', 'cancel'] as const)(
  'rolls back %s during a multi-page import and reopens the previous revision',
  async (fault) => {
    const source = await createWorkbook({ sheets: [{ name: 'Source', rows: 2048, columns: 2 }] });
    const target = await createWorkbook({ sheets: [{ name: 'Target', rows: 100, columns: 2 }] });
    let reopened;
    let unsubscribe = () => {};
    let spy: ReturnType<typeof vi.spyOn> | undefined;
    try {
      await source
        .getSheets()[0]
        .range('A1:B2048')
        .setValues(Array.from({ length: 2048 }, (_, i) => [i, `row-${i}`]));
      const input = await source.toJSON();
      await target
        .getSheets()[0]
        .range('A1')
        .setValues([['keep']]);
      const before = await target.toJSON();
      if (fault === 'quota') {
        const write = Database.prototype.write;
        let attempts = 0;
        spy = vi.spyOn(Database.prototype, 'write').mockImplementation(async function (
          this: Database,
          store,
          puts,
          deletes,
        ) {
          if (
            store === 'records' &&
            String(puts[0]?.[0]).startsWith(`${target.id}/`) &&
            ++attempts === 3
          )
            throw new DOMException('Injected failure', 'QuotaExceededError');
          return write.call(this, store, puts, deletes);
        });
      } else
        unsubscribe = target.onProgress((progress) => {
          if (progress.phase === 'import-store' && progress.completed === 2) target.cancel();
        });
      await expect(target.importJSON(input)).rejects.toThrow();
      spy?.mockRestore();
      unsubscribe();
      expect(await target.toJSON()).toEqual(before);
      const db = await Database.open();
      try {
        const scratch = [];
        for await (const batch of db.scan(
          'scratch',
          IDBKeyRange.bound(`${target.id}/`, `${target.id}/\uffff`),
        ))
          scratch.push(...batch);
        expect(scratch).toEqual([]);
      } finally {
        db.close();
      }
      const id = target.id;
      await target.close();
      reopened = await openWorkbook(id);
      expect(await reopened.getSheets()[0].range('A1').getValues()).toEqual([['keep']]);
      await reopened.undo();
      expect(await reopened.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
      await reopened.importJSON(input);
      expect(await reopened.getSheets()[0].range('A2048:B2048').getValues()).toEqual([
        [2047, 'row-2047'],
      ]);
    } finally {
      spy?.mockRestore();
      unsubscribe();
      await source.close();
      await target.close();
      await reopened?.close();
    }
  },
);

it('reports bounded source-stage and scratch-write totals for streamed imports', async () => {
  const source = await createWorkbook({ sheets: [{ name: 'Source', rows: 130, columns: 2 }] });
  const target = await createWorkbook();
  try {
    await source
      .getSheets()[0]
      .range('A1:B130')
      .setValues(Array.from({ length: 130 }, (_, i) => [i, `row-${i}`]));
    const events: { phase: string; completed: number; elapsedMs?: number }[] = [];
    const stop = target.onProgress((event) => events.push(event));
    try {
      await target.importJSON(new Blob([JSON.stringify(await source.toJSON())]));
    } finally {
      stop();
    }
    const stage = events.find((event) => event.phase === 'import-source-stage');
    const write = events.find((event) => event.phase === 'import-scratch-write');
    expect(stage?.completed).toBeGreaterThan(1);
    expect(write?.completed).toBe(stage?.completed);
    expect(write?.elapsedMs).toBeGreaterThanOrEqual(0);
    expect(stage?.elapsedMs).toBeGreaterThanOrEqual(write!.elapsedMs!);
    expect(await target.getSheets()[0].range('A130:B130').getValues()).toEqual([[129, 'row-129']]);
  } finally {
    await source.close();
    await target.close();
  }
});
