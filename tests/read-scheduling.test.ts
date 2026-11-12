import { expect, it, vi } from 'vitest';
import { Engine } from '../packages/core/src/engine.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { Database } from '../packages/core/src/storage.js';

it('yields according to visited cells, not the sheet-wide coordinate', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 2048, columns: 3 }] });
  const db = await Database.open(`read-scheduling-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId);
  await engine.open(snapshot);
  const id = snapshot.sheets[0].id;
  const timer = vi.spyOn(globalThis, 'setTimeout');
  try {
    await engine.read(id, { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 }, false, false);
    expect(timer).not.toHaveBeenCalled();
    timer.mockClear();
    const result = await engine.read(
      id,
      { startRow: 0, endRow: 2048, startColumn: 1, endColumn: 2 },
      true,
      false,
    );
    expect(Object.keys(result.calculated)).toHaveLength(2048);
    expect(Object.values(result.calculated).every((value) => value === null)).toBe(true);
    expect(timer).toHaveBeenCalledTimes(2);
  } finally {
    timer.mockRestore();
    engine.close();
  }
});

it('honors cancellation delivered at the final scheduling checkpoint', async () => {
  const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 1024, columns: 1 }] });
  const db = await Database.open(`read-cancel-${crypto.randomUUID()}`);
  const engine = new Engine(db, snapshot.workbookId);
  await engine.open(snapshot);
  const original = globalThis.setTimeout;
  const timer = vi.spyOn(globalThis, 'setTimeout').mockImplementationOnce(((callback: () => void) =>
    original(() => {
      engine.cancel();
      callback();
    }, 0)) as typeof setTimeout);
  try {
    await expect(
      engine.read(
        snapshot.sheets[0].id,
        { startRow: 0, endRow: 1024, startColumn: 0, endColumn: 1 },
        true,
        false,
      ),
    ).rejects.toMatchObject({ code: 'ABORTED' });
    engine.begin();
    const retry = await engine.read(
      snapshot.sheets[0].id,
      { startRow: 0, endRow: 1024, startColumn: 0, endColumn: 1 },
      true,
      false,
    );
    expect(Object.keys(retry.calculated)).toHaveLength(1024);
  } finally {
    timer.mockRestore();
    engine.close();
  }
});
