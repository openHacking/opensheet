import { expect, it, vi } from 'vitest';
import { verificationYield } from '../packages/core/src/verification-yield.js';

it('yields after its task budget and observes cancellation before resuming', async () => {
  vi.useFakeTimers();
  try {
    const controller = new AbortController();
    const checkpoint = verificationYield(controller.signal);
    await checkpoint();
    vi.advanceTimersByTime(9);
    const pending = checkpoint();
    const rejected = expect(pending).rejects.toMatchObject({ name: 'AbortError' });
    controller.abort();
    await vi.runAllTimersAsync();
    await rejected;
  } finally {
    vi.useRealTimers();
  }
});
