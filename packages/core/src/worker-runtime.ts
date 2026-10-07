import { Engine } from './engine.js';
import { encodeWorkbookFile } from './file-codec.js';
import { Database } from './storage.js';
import type { Command, Rect } from './types.js';
export type Message = { id: number; method: string; args: any[] };
export function serveWorker(scope: {
  onmessage: ((event: MessageEvent<Message>) => void) | null;
  postMessage(value: unknown): void;
}) {
  let engine: Engine | undefined;
  let queue = Promise.resolve();
  let sentMetadata: string | undefined;
  let compactTimer: ReturnType<typeof setTimeout> | undefined;
  scope.onmessage = (event) => {
    const { id, method, args } = event.data;
    clearTimeout(compactTimer);
    if (method === 'cancel') {
      engine?.cancel();
      return;
    }
    queue = queue.then(async () => {
      try {
        let result: unknown;
        if (method === 'open') {
          const [bookId, snapshot, options] = args as [
            string,
            import('./file-codec.js').WorkbookFile | undefined,
            { database?: string; budgetBytes?: number; cacheBytes?: number },
          ];
          engine = new Engine(
            await Database.open(options.database),
            bookId,
            options.cacheBytes,
            (progress) => scope.postMessage({ progress, requestId: id }),
          );
          result = {
            view: await engine.open(snapshot, options.budgetBytes),
            writer: engine.isWriter,
          };
        } else {
          if (!engine) throw new Error('Worker is not initialized');
          engine.begin();
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
            case 'scanCells':
              result = await engine.scanCells(args[0], args[1]);
              break;
            case 'read':
              result = await engine.read(args[0] as string, args[1] as Rect, args[2]);
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
              result = await engine.search(args[0], args[1], args[2]);
              break;
            case 'compact':
              result = await engine.compact(args[0]);
              break;
            case 'stats': {
              const view = engine.view();
              result = { bytes: view.bytes, cacheBytes: view.cacheBytes, budgetBytes: view.budget };
              break;
            }
            case 'reload':
              result = await engine.reload();
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
            if (sentMetadata !== view.metadataId) {
              view.file = encodeWorkbookFile(snapshot);
              view.styleIds = Object.keys(snapshot.styles);
            }
            sentMetadata = view.metadataId;
          }
        }
        scope.postMessage({ id, result });
        if (method !== 'close' && method !== 'clear') {
          const compact = () => {
            queue = queue.then(async () => {
              try {
                const storage = await engine?.compact(1);
                if (storage) scope.postMessage({ storage });
                if (engine?.hasPendingCompaction) compactTimer = setTimeout(compact, 100);
              } catch {
                /* Keep raw records if compression fails. */
              }
            });
          };
          compactTimer = setTimeout(compact, 1000);
        }
      } catch (error) {
        scope.postMessage({
          id,
          error: {
            message: error instanceof Error ? error.message : String(error),
            code:
              typeof (error as { code?: unknown }).code === 'string'
                ? (error as { code: string }).code
                : (error as Error).name,
          },
        });
      }
    });
  };
}
