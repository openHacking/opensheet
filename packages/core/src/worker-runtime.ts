import { Engine } from './engine.js';
import { encodeWorkbookFile } from './file-codec.js';
import { Database } from './storage.js';
import type { Command, Rect } from './types.js';
export type Message = { id: number; method: string; args: any[]; metadataId?: string };
export function serveWorker(scope: {
  onmessage: ((event: MessageEvent<Message>) => void) | null;
  postMessage(value: unknown): void;
}) {
  let engine: Engine | undefined;
  let queue = Promise.resolve();
  let operationTiming = false;
  let activeMethod = '';
  let waitingTasks = 0;
  let compactTimer: ReturnType<typeof setTimeout> | undefined;
  let compactQueued = false;
  const requests = new Map<number, string>();
  const cancelledRequests = new Set<number>();
  let activeRequest: number | undefined;
  function scheduleCompact(delay: number) {
    clearTimeout(compactTimer);
    compactTimer = setTimeout(() => {
      compactTimer = undefined;
      if (compactQueued) return;
      compactQueued = true;
      const traced = operationTiming;
      if (traced) waitingTasks++;
      queue = queue.then(async () => {
        if (traced) {
          waitingTasks--;
          activeMethod = 'compact';
        }
        let again = false;
        try {
          const storage = await engine?.compact(1);
          if (storage) scope.postMessage({ storage });
          again = !!engine?.hasPendingCompaction;
        } catch {
          /* Keep raw records if compression fails. */
        } finally {
          compactQueued = false;
          if (traced) activeMethod = '';
        }
        if (again) scheduleCompact(100);
      });
    }, delay);
  }
  scope.onmessage = (event) => {
    const { id, method, args, metadataId } = event.data;
    clearTimeout(compactTimer);
    compactTimer = undefined;
    if (method === 'cancel') {
      engine?.cancel();
      return;
    }
    if (method === 'cancelRequest') {
      const target = args[0];
      if (['search', 'navigateData', 'readCells'].includes(requests.get(target) ?? '')) {
        if (activeRequest === target) engine?.cancel();
        else cancelledRequests.add(target);
      }
      return;
    }
    requests.set(id, method);
    const queuedAt = operationTiming ? performance.now() : undefined;
    const activeAtArrival = queuedAt === undefined ? '' : activeMethod;
    const queuedAhead = queuedAt === undefined ? 0 : waitingTasks;
    if (queuedAt !== undefined) waitingTasks++;
    queue = queue.then(async () => {
      if (queuedAt !== undefined) {
        waitingTasks--;
        activeMethod = method;
      }
      const startedAt = queuedAt === undefined ? undefined : performance.now();
      const readTiming =
        startedAt !== undefined && method === 'read'
          ? {
              cacheHits: 0,
              cacheMisses: 0,
              absentPages: 0,
              storageMs: 0,
              decodeMs: 0,
              inflateMs: 0,
              pageDecodeMs: 0,
            }
          : undefined;
      const timing = () =>
        startedAt === undefined || queuedAt === undefined
          ? undefined
          : {
              queueMs: startedAt - queuedAt,
              executionMs: performance.now() - startedAt,
              activeAtArrival,
              queuedAhead,
              ...(readTiming ? { read: readTiming } : {}),
            };
      try {
        activeRequest = id;
        if (cancelledRequests.delete(id))
          throw Object.assign(new Error('Operation cancelled'), { code: 'ABORTED' });
        let result: unknown;
        if (method === 'open') {
          const [bookId, snapshot, options] = args as [
            string,
            import('./file-codec.js').WorkbookFile | undefined,
            {
              database?: string;
              budgetBytes?: number;
              cacheBytes?: number;
              preserveOrphans?: boolean;
              operationTiming?: boolean;
            },
          ];
          operationTiming = !!options.operationTiming;
          engine = new Engine(
            await Database.open(options.database),
            bookId,
            options.cacheBytes,
            (progress) => scope.postMessage({ progress, requestId: id }),
          );
          result = {
            view: await engine.open(snapshot, options.budgetBytes, options.preserveOrphans),
            writer: engine.isWriter,
          };
        } else {
          if (!engine) throw new Error('Worker is not initialized');
          if (method !== 'close') engine.begin();
          switch (method) {
            case 'exportIndex':
              result = engine.exportIndex(args[0], args[1], args[2], args[3]);
              break;
            case 'exportStyles':
              result = engine.exportStyles(args[0]);
              break;
            case 'exportBlock':
              result = await engine.exportBlock(args[0], args[1]);
              break;
            case 'readCells':
              result = await engine.readCells(args[0], args[1], args[2]);
              break;
            case 'navigateData':
              result = await engine.navigateData(args[0], args[1], args[2], args[3], args[4]);
              break;
            case 'scanCells':
              result = await engine.scanCells(args[0], args[1]);
              break;
            case 'read':
              result = await engine.read(
                args[0] as string,
                args[1] as Rect,
                args[2],
                true,
                readTiming,
              );
              break;
            case 'execute':
              result = await engine.execute(args[0] as Command[], args[1]);
              break;
            case 'history':
              result = await engine.history(args[0]);
              break;
            case 'generate':
              result = await engine.generate(args[0], args[1]);
              break;
            case 'importBinary':
              result = await engine.importBinary(args[0]);
              break;
            case 'importJSON':
              result = await engine.importJSON(args[0]);
              break;
            case 'replace':
              result = await engine.replace(args[0], args[1]);
              break;
            case 'search':
              result = await engine.search(args[0], args[1], args[2], args[3], args[4]);
              break;
            case 'compact':
              result = await engine.compact(args[0]);
              break;
            case 'stats': {
              const view = engine.view();
              result = { bytes: view.bytes, cacheBytes: view.cacheBytes, budgetBytes: view.budget };
              break;
            }
            case 'tryAcquireWriteLock':
              result = await engine.tryAcquireWriteLock();
              break;
            case 'reload':
              result = await engine.reload();
              break;
            case 'reloadIfChanged':
              result = await engine.reloadIfChanged();
              break;
            case 'headExists':
              result = await engine.headExists();
              break;
            case 'clear':
              result = await engine.clear();
              break;
            case 'close':
              engine.close();
              result = null;
              break;
            default:
              throw new Error(`Unknown worker method: ${method}`);
          }
        }
        if (result && typeof result === 'object' && ('view' in result || 'snapshot' in result)) {
          const container = result as any,
            view = container.view ?? container;
          if (view.snapshot) {
            const snapshot = view.snapshot;
            delete view.snapshot;
            if (metadataId !== view.metadataId) {
              view.file = encodeWorkbookFile(snapshot);
              view.styleIds = Object.keys(snapshot.styles);
            }
          }
        }
        scope.postMessage({ id, result, ...(startedAt === undefined ? {} : { timing: timing() }) });
        if (method !== 'close' && method !== 'clear') scheduleCompact(1000);
        else {
          clearTimeout(compactTimer);
          compactTimer = undefined;
        }
      } catch (error) {
        scope.postMessage({
          id,
          ...(startedAt === undefined ? {} : { timing: timing() }),
          error: {
            message: error instanceof Error ? error.message : String(error),
            code:
              typeof (error as { code?: unknown }).code === 'string'
                ? (error as { code: string }).code
                : (error as Error).name,
          },
        });
      } finally {
        requests.delete(id);
        cancelledRequests.delete(id);
        activeRequest = undefined;
        if (startedAt !== undefined) activeMethod = '';
      }
    });
  };
}
