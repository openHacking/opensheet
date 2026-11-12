type Summary = { workbookId: string; sheets: number; pages: number; cells: number };

/** Verify in a disposable Worker so large backup parsing does not block the UI. */
export function verifyBackupInWorker(file: File, signal: AbortSignal): Promise<Summary> {
  signal.throwIfAborted();
  const worker = new Worker(new URL('./backup-verifier.worker.ts', import.meta.url), {
    type: 'module',
  });
  return new Promise((resolve, reject) => {
    let finished = false;
    const finish = (error?: unknown, summary?: Summary) => {
      if (finished) return;
      finished = true;
      signal.removeEventListener('abort', abort);
      worker.terminate();
      if (error !== undefined) reject(error);
      else resolve(summary!);
    };
    const abort = () => finish(signal.reason);
    worker.onmessage = (event: MessageEvent<{ summary?: Summary; error?: string }>) => {
      if (event.data.error !== undefined) finish(new Error(event.data.error));
      else if (event.data.summary) finish(undefined, event.data.summary);
      else finish(new Error('Invalid verification response'));
    };
    worker.onerror = (event) =>
      finish(new Error(event.message || 'Backup verification worker failed'));
    worker.onmessageerror = () => finish(new Error('Invalid verification response'));
    signal.addEventListener('abort', abort, { once: true });
    if (signal.aborted) abort();
    else {
      try {
        worker.postMessage({ file, json: /\.json$/i.test(file.name) });
      } catch (error) {
        finish(error);
      }
    }
  });
}
