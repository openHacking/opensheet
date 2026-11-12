/** Yield between completed pages when verification has occupied the current task budget. */
export function verificationYield(signal?: AbortSignal) {
  let deadline = performance.now() + 8;
  return async () => {
    signal?.throwIfAborted();
    if (performance.now() < deadline) return;
    await new Promise<void>((resolve) => setTimeout(resolve, 0));
    signal?.throwIfAborted();
    deadline = performance.now() + 8;
  };
}

/** Stop waiting on a pending operation; the operation itself remains caller-owned. */
export function abortable<T>(operation: Promise<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return operation;
  return new Promise<T>((resolve, reject) => {
    const abort = () => {
      signal.removeEventListener('abort', abort);
      reject(signal.reason);
    };
    signal.addEventListener('abort', abort, { once: true });
    operation.then(
      (value) => {
        signal.removeEventListener('abort', abort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener('abort', abort);
        reject(error);
      },
    );
    if (signal.aborted) abort();
  });
}
