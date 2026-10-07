import 'fake-indexeddb/auto';
import { serveWorker, type Message } from '../packages/core/src/worker-runtime.js';
/** Contract tests run the exact worker dispatcher against IndexedDB, browser E2E uses real Workers. */
class TestWorker {
  onmessage?: (event: MessageEvent) => void;
  onerror?: (event: ErrorEvent) => void;
  onmessageerror?: () => void;
  private alive = true;
  private scope = {
    onmessage: null as ((event: MessageEvent<Message>) => void) | null,
    postMessage: (value: unknown) => {
      queueMicrotask(() => {
        if (this.alive) this.onmessage?.({ data: structuredClone(value) } as MessageEvent);
      });
    },
  };
  constructor() {
    serveWorker(this.scope);
  }
  postMessage(value: Message) {
    const data = structuredClone(value);
    queueMicrotask(() => {
      if (this.alive) this.scope.onmessage?.({ data } as MessageEvent<Message>);
    });
  }
  terminate() {
    this.scope.onmessage?.({
      data: { id: -1, method: 'close', args: [] },
    } as unknown as MessageEvent<Message>);
    this.alive = false;
  }
}
Object.assign(globalThis, { Worker: TestWorker });
