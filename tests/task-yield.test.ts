import { afterEach, expect, it, vi } from 'vitest';
import { yieldToTask } from '../packages/core/src/task-yield.js';

afterEach(() => vi.unstubAllGlobals());

it('closes both message ports after yielding', async () => {
  const closeFirst = vi.fn(),
    closeSecond = vi.fn();
  class Channel {
    port1 = { onmessage: null as (() => void) | null, close: closeFirst };
    port2 = {
      close: closeSecond,
      postMessage: () => queueMicrotask(() => this.port1.onmessage?.()),
    };
  }
  vi.stubGlobal('MessageChannel', Channel);
  await yieldToTask();
  expect(closeFirst).toHaveBeenCalledOnce();
  expect(closeSecond).toHaveBeenCalledOnce();
});

it('uses a timer when MessageChannel is unavailable', async () => {
  vi.stubGlobal('MessageChannel', undefined);
  const timer = vi.spyOn(globalThis, 'setTimeout');
  try {
    await yieldToTask();
    expect(timer).toHaveBeenCalled();
  } finally {
    timer.mockRestore();
  }
});
