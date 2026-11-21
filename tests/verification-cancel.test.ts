import { expect, it, vi } from 'vitest';
import { compareBinaryBackup, createWorkbook, verifyBinary, verifyJSON } from '@opensheetjs/core';
import { unpackPage, writeBinary } from '../packages/core/src/binary-file.js';

it('does not release or cancel an internally owned completed decompression stream', async () => {
  let released = false;
  let cancelled = false;
  class PassthroughDecompression {
    private stream = new TransformStream<Uint8Array, Uint8Array>({
      transform(chunk, controller) {
        controller.enqueue(chunk);
      },
    });
    readable = this.stream.readable;
    writable = this.stream.writable;
    constructor() {
      const getReader = this.readable.getReader.bind(this.readable);
      this.readable.getReader = (() => {
        const reader = getReader();
        reader.releaseLock = () => {
          released = true;
          throw new Error('releaseLock must not be used');
        };
        const cancel = reader.cancel.bind(reader);
        reader.cancel = (...args) => {
          cancelled = true;
          return cancel(...args);
        };
        return reader;
      }) as typeof this.readable.getReader;
    }
  }
  vi.stubGlobal('DecompressionStream', PassthroughDecompression);
  try {
    expect(await unpackPage(new Uint8Array([1, 2, 3]), 1, 3)).toEqual(new Uint8Array([1, 2, 3]));
    expect(released).toBe(false);
    expect(cancelled).toBe(false);
  } finally {
    vi.unstubAllGlobals();
  }
});

it('reports a compressed page overflow without waiting for stalled stream cancellation', async () => {
  let cancelled = 0;
  class OverflowDecompression {
    private stream = new TransformStream<Uint8Array, Uint8Array>({
      transform(chunk, controller) {
        controller.enqueue(chunk);
      },
    });
    readable = this.stream.readable;
    writable = this.stream.writable;
    constructor() {
      const getReader = this.readable.getReader.bind(this.readable);
      this.readable.getReader = (() => {
        const reader = getReader();
        reader.cancel = () => {
          cancelled++;
          return new Promise<void>(() => {});
        };
        return reader;
      }) as typeof this.readable.getReader;
    }
  }
  vi.stubGlobal('DecompressionStream', OverflowDecompression);
  try {
    const outcome = await Promise.race([
      unpackPage(new Uint8Array([1, 2]), 1, 1).then(
        () => ({ code: 'unexpected-success' }),
        (error: { code?: string }) => ({ code: error.code }),
      ),
      new Promise<{ code: string }>((resolve) =>
        setTimeout(() => resolve({ code: 'timeout' }), 1000),
      ),
    ]);
    expect(outcome).toEqual({ code: 'LIMIT_EXCEEDED' });
    expect(cancelled).toBe(1);
  } finally {
    vi.unstubAllGlobals();
  }
});

for (const verify of [verifyJSON, verifyBinary]) {
  it('cancels verification while input and stream cleanup are stalled', async () => {
    let cancelled = false;
    const stream = new ReadableStream<Uint8Array<ArrayBuffer>>({
      pull: () => new Promise<void>(() => {}),
      cancel: () => {
        cancelled = true;
        return new Promise<void>(() => {});
      },
    });
    const blob = new Blob(['placeholder']);
    const mock = vi.spyOn(blob, 'stream').mockReturnValue(stream);
    const controller = new AbortController();
    try {
      const result = verify(blob, { signal: controller.signal });
      await Promise.resolve();
      controller.abort();
      await expect(result).rejects.toMatchObject({ name: 'AbortError' });
      expect(cancelled).toBe(true);
      expect(stream.locked).toBe(false);
    } finally {
      mock.mockRestore();
    }
  }, 1000);
}

it('cancels while compressed page decoding is stalled', async () => {
  let started!: () => void;
  const decoding = new Promise<void>((resolve) => {
    started = resolve;
  });

  class StalledDecompression {
    private stream = new TransformStream<Uint8Array, Uint8Array>({
      transform() {
        started();
        return new Promise<void>(() => {});
      },
    });
    readable = this.stream.readable;
    writable = this.stream.writable;
  }
  vi.stubGlobal('DecompressionStream', StalledDecompression);
  const controller = new AbortController();
  try {
    const result = unpackPage(new Uint8Array([1]), 1, 1, undefined, controller.signal);
    await decoding;
    controller.abort();
    await expect(result).rejects.toMatchObject({ name: 'AbortError' });
  } finally {
    vi.unstubAllGlobals();
  }
}, 1000);

it('retains the explicit abort reason if decompression fails after cancellation', async () => {
  let started!: () => void;
  let rejectDecode!: (reason: Error) => void;
  const decoding = new Promise<void>((resolve) => {
    started = resolve;
  });
  class FailingDecompression {
    private stream = new TransformStream<Uint8Array, Uint8Array>({
      transform() {
        started();
        return new Promise<void>((_resolve, reject) => {
          rejectDecode = reject;
        });
      },
    });
    readable = this.stream.readable;
    writable = this.stream.writable;
  }
  vi.stubGlobal('DecompressionStream', FailingDecompression);
  const controller = new AbortController(),
    reason = new Error('explicit cancellation');
  try {
    const result = unpackPage(new Uint8Array([1]), 1, 1, undefined, controller.signal);
    await decoding;
    controller.abort(reason);
    rejectDecode(new Error('late decompression failure'));
    await expect(result).rejects.toBe(reason);
  } finally {
    vi.unstubAllGlobals();
  }
}, 1000);

it('passes cancellation through binary verification into page decompression', async () => {
  const book = await createWorkbook();
  try {
    const metadata = await book.toJSON();
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes: new Uint8Array(4096) };
    }
    const parts: Uint8Array[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    expect(new DataView(parts[2].buffer, parts[2].byteOffset).getUint32(12, true)).toBe(1);
    const file = new Blob(parts as BlobPart[]);
    let started!: () => void;
    const decoding = new Promise<void>((resolve) => {
      started = resolve;
    });
    class StalledDecompression {
      private stream = new TransformStream<Uint8Array, Uint8Array>({
        transform() {
          started();
          return new Promise<void>(() => {});
        },
      });
      readable = this.stream.readable;
      writable = this.stream.writable;
    }
    vi.stubGlobal('DecompressionStream', StalledDecompression);
    const controller = new AbortController();
    try {
      const result = verifyBinary(file, { signal: controller.signal });
      await decoding;
      controller.abort();
      await expect(result).rejects.toMatchObject({ name: 'AbortError' });
    } finally {
      vi.unstubAllGlobals();
    }
  } finally {
    await book.deleteStorage();
    await book.close();
  }
}, 2000);

it('cancels a stalled backup comparison read without waiting for stream cleanup', async () => {
  const book = await createWorkbook();
  try {
    const parts: Uint8Array[] = [];
    for await (const part of book.streamBinary()) parts.push(part);
    const blob = new Blob(parts as BlobPart[]);
    const original = blob.stream.bind(blob);
    let calls = 0,
      cancelled = false;
    let started!: () => void;
    const reading = new Promise<void>((resolve) => {
      started = resolve;
    });
    const stalled = new ReadableStream<Uint8Array<ArrayBuffer>>({
      pull: () => new Promise<void>(() => {}),
      cancel: () => {
        cancelled = true;
        return new Promise<void>(() => {});
      },
    });
    const mock = vi.spyOn(blob, 'stream').mockImplementation(() => {
      if (++calls === 1) return original();
      started();
      return stalled;
    });
    const controller = new AbortController();
    try {
      const result = compareBinaryBackup(book, blob, { signal: controller.signal });
      await reading;
      controller.abort();
      await expect(result).rejects.toMatchObject({ name: 'AbortError' });
      expect(cancelled).toBe(true);
      expect(stalled.locked).toBe(false);
    } finally {
      mock.mockRestore();
    }
  } finally {
    await book.deleteStorage();
  }
}, 2000);

it('returns promptly when cancellation interrupts a pending export chunk', async () => {
  const book = await createWorkbook();
  let release!: () => void;
  const pending = new Promise<void>((resolve) => {
    release = resolve;
  });
  try {
    const parts: Uint8Array[] = [];
    for await (const part of book.streamBinary()) parts.push(part);
    const blob = new Blob(parts as BlobPart[]);
    let started!: () => void;
    const waiting = new Promise<void>((resolve) => {
      started = resolve;
    });
    const original = book.streamBinary.bind(book);
    book.streamBinary = async function* () {
      started();
      await pending;
      yield new Uint8Array([0]);
    };
    const controller = new AbortController();
    try {
      const result = compareBinaryBackup(book, blob, { signal: controller.signal });
      await waiting;
      controller.abort();
      await expect(result).rejects.toMatchObject({ name: 'AbortError' });
    } finally {
      release();
      book.streamBinary = original;
    }
  } finally {
    release();
    await book.deleteStorage();
  }
}, 2000);
