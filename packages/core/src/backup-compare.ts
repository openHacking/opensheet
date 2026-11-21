import { verifyBinary } from './binary-verify.js';
import { assert } from './types.js';
import { abortable, verificationYield } from './verification-yield.js';
import type { Workbook } from './workbook.js';

/** Compare a validated backup with a fresh export of one revision, byte for byte. */
export async function compareBinaryBackup(
  book: Workbook,
  input: Blob,
  options: { signal?: AbortSignal } = {},
) {
  assert(input instanceof Blob, 'INVALID_ARGUMENT', 'Backup comparison requires a Blob');
  const signal = options.signal;
  signal?.throwIfAborted();
  const revision = book.revision;
  const stable = () =>
    assert(
      book.revision === revision,
      'REVISION_CONFLICT',
      'Workbook changed during backup comparison',
    );
  await verifyBinary(input, options);
  stable();
  const expected = book.streamBinary(options);
  const actual = input.stream().getReader();
  const abort = () => {
    void actual.cancel().catch(() => {});
  };
  signal?.addEventListener('abort', abort, { once: true });
  const checkpoint = verificationYield(signal);
  let left: Uint8Array = new Uint8Array(0),
    right: Uint8Array = new Uint8Array(0),
    l = 0,
    r = 0;
  let leftDone = false,
    rightDone = false;
  try {
    while (true) {
      signal?.throwIfAborted();
      stable();
      if (l === left.length && !leftDone) {
        const next = await abortable(expected.next(), signal);
        leftDone = !!next.done;
        left = next.value ?? new Uint8Array(0);
        l = 0;
      }
      if (r === right.length && !rightDone) {
        const next = await actual.read();
        rightDone = next.done;
        right = next.value ?? new Uint8Array(0);
        r = 0;
      }
      signal?.throwIfAborted();
      stable();
      if (leftDone || rightDone) return { matches: leftDone && rightDone, revision };
      const size = Math.min(left.length - l, right.length - r, 65536);
      for (let i = 0; i < size; i++)
        if (left[l + i] !== right[r + i]) return { matches: false, revision };
      l += size;
      r += size;
      await checkpoint();
    }
  } finally {
    signal?.removeEventListener('abort', abort);
    // A caller-provided Blob stream can leave cancellation pending forever.
    void actual.cancel().catch(() => {});
    actual.releaseLock();
    const closing = expected.return(undefined);
    if (signal?.aborted) void closing.catch(() => {});
    else await closing;
  }
}
