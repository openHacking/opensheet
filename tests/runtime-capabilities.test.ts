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

  it('rejects opaque browser origins before allocating a channel or Worker', () => {
    const worker = vi.fn();
    const channel = vi.fn();
    vi.stubGlobal('window', {});
    vi.stubGlobal('origin', 'null');
    vi.stubGlobal('Worker', worker);
    vi.stubGlobal('BroadcastChannel', channel);
    expect(() => new Workbook()).toThrow(
      expect.objectContaining({
        code: 'STORAGE_UNAVAILABLE',
        message: expect.stringContaining('opaque origin'),
      }),
    );
    expect(worker).not.toHaveBeenCalled();
    expect(channel).not.toHaveBeenCalled();
    const failure = new Error('injected runtime reached');
    expect(
      () =>
        new Workbook(undefined, {
          workerFactory: () => {
            throw failure;
          },
        }),
    ).toThrow(failure);
  });

  it('reports denied Worker property access before attempting initialization', () => {
    const original = Object.getOwnPropertyDescriptor(globalThis, 'Worker');
    Object.defineProperty(globalThis, 'Worker', {
      configurable: true,
      get: () => {
        throw new DOMException('Blocked', 'SecurityError');
      },
    });
    try {
      expect(() => new Workbook()).toThrow(expect.objectContaining({ code: 'WORKER_UNAVAILABLE' }));
    } finally {
      if (original) Object.defineProperty(globalThis, 'Worker', original);
      else Reflect.deleteProperty(globalThis, 'Worker');
    }
  });

  it('rejects a non-constructor Worker before attempting initialization', () => {
    vi.stubGlobal('Worker', {});
    expect(() => new Workbook()).toThrow(expect.objectContaining({ code: 'WORKER_UNAVAILABLE' }));
  });

  it.each(['SecurityError', 'NotAllowedError'])(
    'reports %s when the browser denies Worker construction',
    (name) => {
      vi.stubGlobal(
        'Worker',
        class {
          constructor() {
            throw new DOMException('Blocked', name);
          }
        },
      );
      expect(() => new Workbook()).toThrow(
        expect.objectContaining({
          code: 'WORKER_UNAVAILABLE',
          message: expect.stringContaining('security policy'),
        }),
      );
    },
  );

  it('preserves unexpected Worker construction and injected factory errors', () => {
    const browserError = new TypeError('Worker bug');
    vi.stubGlobal(
      'Worker',
      class {
        constructor() {
          throw browserError;
        }
      },
    );
    expect(() => new Workbook()).toThrow(browserError);
    const factoryError = new DOMException('Host policy', 'SecurityError');
    expect(
      () =>
        new Workbook(undefined, {
          workerFactory: () => {
            throw factoryError;
          },
        }),
    ).toThrow(factoryError);
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

  it.each(['SecurityError', 'NotAllowedError', 'TypeError'])(
    'rejects %s from BroadcastChannel construction before creating the Worker',
    (name) => {
      const worker = vi.fn();
      const failure =
        name === 'TypeError' ? new TypeError('Unexpected') : new DOMException('Blocked', name);
      vi.stubGlobal('Worker', worker);
      vi.stubGlobal(
        'BroadcastChannel',
        class {
          constructor() {
            throw failure;
          }
        },
      );
      if (name === 'TypeError') expect(() => new Workbook()).toThrow(failure);
      else
        expect(() => new Workbook()).toThrow(
          expect.objectContaining({ code: 'BROADCAST_CHANNEL_UNAVAILABLE' }),
        );
      expect(worker).not.toHaveBeenCalled();
    },
  );

  it.each(['browser', 'factory'] as const)(
    'closes the channel if %s Worker construction fails, preserving the startup error',
    (kind) => {
      const failure = new Error('Worker failed');
      const close = vi.fn(() => {
        throw new Error('Close failed');
      });
      vi.stubGlobal(
        'BroadcastChannel',
        class {
          close = close;
        },
      );
      vi.stubGlobal(
        'Worker',
        class {
          constructor() {
            if (kind === 'browser') throw failure;
          }
        },
      );
      const factory =
        kind === 'factory'
          ? () => {
              throw failure;
            }
          : undefined;
      expect(() => new Workbook(undefined, { workerFactory: factory })).toThrow(failure);
      expect(close).toHaveBeenCalledOnce();
    },
  );

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
