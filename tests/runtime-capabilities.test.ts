import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, Workbook } from '@opensheetjs/core';
import { Database } from '../packages/core/src/storage.js';

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe('runtime prerequisites', () => {
  it.each([
    ['Worker', 'WORKER_UNAVAILABLE'],
    ['BroadcastChannel', 'BROADCAST_CHANNEL_UNAVAILABLE'],
    ['indexedDB', 'STORAGE_UNAVAILABLE'],
    ['navigator', 'WEB_LOCKS_UNAVAILABLE'],
  ])('rejects missing %s before allocating a Worker', (capability, code) => {
    const worker = vi.fn();
    vi.stubGlobal('window', {});
    vi.stubGlobal('Worker', worker);
    vi.stubGlobal('navigator', { locks: { request: vi.fn() } });
    vi.stubGlobal(capability, undefined);
    expect(() => new Workbook()).toThrow(expect.objectContaining({ code }));
    expect(worker).not.toHaveBeenCalled();
  });

  it('allows an injected storage runtime without browser Web Locks', async () => {
    const WorkerImplementation = Worker;
    const factory = vi.fn(() => new WorkerImplementation(''));
    // The test Worker runs in-process, so leave its IndexedDB implementation available.
    vi.stubGlobal('navigator', {});
    const book = await createWorkbook({}, { workerFactory: factory });
    expect(factory).toHaveBeenCalledOnce();
    await book.close();
  });

  it('reports denied IndexedDB property access before allocating a Worker', () => {
    const original = Object.getOwnPropertyDescriptor(globalThis, 'indexedDB');
    const worker = vi.fn();
    vi.stubGlobal('window', {});
    vi.stubGlobal('Worker', worker);
    Object.defineProperty(globalThis, 'indexedDB', {
      configurable: true,
      get: () => {
        throw new DOMException('Blocked', 'SecurityError');
      },
    });
    try {
      expect(() => new Workbook()).toThrow(
        expect.objectContaining({ code: 'STORAGE_UNAVAILABLE' }),
      );
      expect(worker).not.toHaveBeenCalled();
    } finally {
      if (original) Object.defineProperty(globalThis, 'indexedDB', original);
      else Reflect.deleteProperty(globalThis, 'indexedDB');
    }
  });

  it('reports denied Web Locks property access before allocating a Worker', () => {
    const worker = vi.fn();
    vi.stubGlobal('window', {});
    vi.stubGlobal('Worker', worker);
    vi.stubGlobal('navigator', {
      get locks() {
        throw new DOMException('Blocked', 'SecurityError');
      },
    });
    expect(() => new Workbook()).toThrow(
      expect.objectContaining({ code: 'WEB_LOCKS_UNAVAILABLE' }),
    );
    expect(worker).not.toHaveBeenCalled();
  });

  it('reports denied BroadcastChannel property access before allocating a Worker', () => {
    const original = Object.getOwnPropertyDescriptor(globalThis, 'BroadcastChannel');
    const worker = vi.fn();
    vi.stubGlobal('Worker', worker);
    Object.defineProperty(globalThis, 'BroadcastChannel', {
      configurable: true,
      get: () => {
        throw new DOMException('Blocked', 'SecurityError');
      },
    });
    try {
      expect(() => new Workbook()).toThrow(
        expect.objectContaining({ code: 'BROADCAST_CHANNEL_UNAVAILABLE' }),
      );
      expect(worker).not.toHaveBeenCalled();
    } finally {
      if (original) Object.defineProperty(globalThis, 'BroadcastChannel', original);
      else Reflect.deleteProperty(globalThis, 'BroadcastChannel');
    }
  });

  it('rejects a non-constructor BroadcastChannel before allocating a Worker', () => {
    const worker = vi.fn();
    vi.stubGlobal('Worker', worker);
    vi.stubGlobal('BroadcastChannel', {});
    expect(() => new Workbook()).toThrow(
      expect.objectContaining({ code: 'BROADCAST_CHANNEL_UNAVAILABLE' }),
    );
    expect(worker).not.toHaveBeenCalled();
  });

  it.each(['synchronous', 'asynchronous'] as const)(
    'maps %s IndexedDB security failures without disguising other errors',
    async (mode) => {
      const fail = (name: string) => {
        const error = new DOMException('Injected failure', name);
        vi.spyOn(indexedDB, 'open').mockImplementationOnce(() => {
          if (mode === 'synchronous') throw error;
          const fake = { error, onerror: null as (() => void) | null };
          queueMicrotask(() => fake.onerror?.());
          return fake as unknown as IDBOpenDBRequest;
        });
      };
      fail('SecurityError');
      await expect(Database.open()).rejects.toMatchObject({
        code: 'STORAGE_UNAVAILABLE',
        message: expect.stringContaining('enable site storage'),
      });
      fail('QuotaExceededError');
      await expect(Database.open()).rejects.toMatchObject({ name: 'QuotaExceededError' });
    },
  );
});
