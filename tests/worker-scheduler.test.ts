import { afterEach, expect, test, vi } from 'vitest';
import { createWorkbook } from '../packages/core/src/workbook.js';
import { Engine } from '../packages/core/src/engine.js';

afterEach(() => {
  vi.restoreAllMocks();
  vi.useRealTimers();
});

test('coalesces idle compression while requests are queued', async () => {
  vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout'] });
  let release!: () => void;
  const gate = new Promise<void>((resolve) => {
    release = resolve;
  });
  let count = 0;
  const compact = vi.spyOn(Engine.prototype, 'compact').mockImplementation(async () => {
    count++;
    if (count === 1) await gate;
    return { bytes: 0, cacheBytes: 0 };
  });
  const book = await createWorkbook();
  try {
    await Promise.all(Array.from({ length: 40 }, () => book.storageStats()));
    expect(vi.getTimerCount()).toBe(1);
    await vi.advanceTimersByTimeAsync(1000);
    expect(compact).toHaveBeenCalledTimes(1);
    const pending = Array.from({ length: 40 }, () => book.storageStats());
    await vi.advanceTimersByTimeAsync(5000);
    expect(compact).toHaveBeenCalledTimes(1);
    release();
    await Promise.all(pending);
    expect(vi.getTimerCount()).toBe(1);
    await vi.advanceTimersByTimeAsync(1000);
    expect(compact).toHaveBeenCalledTimes(2);
  } finally {
    release();
    await book.close();
    vi.clearAllTimers();
  }
});
