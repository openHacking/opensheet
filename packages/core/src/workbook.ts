import { freeze } from 'immer';
import { validateSearchArguments } from './search-validation.js';
import { validateNavigationArguments } from './navigation-validation.js';
import {
  encodeWorkbookFile,
  encodeFileCells,
  decodeFileMetadata,
  type WorkbookFile,
} from './file-codec.js';
import { decodeBlock, encodeBlock, type EncodedBlock } from './block-codec.js';
import { writeBinary } from './binary-file.js';
import { validateCommand } from './commands.js';
import { createSheet, createSnapshot, checkBounds } from './model.js';
import { Sheet } from './sheet.js';
import { BLOCK_ROWS, BLOCK_COLUMNS, ByteCache, encodedBytes } from './storage.js';
import { blockAddress, type EngineView, type Progress, type ReadTiming } from './engine.js';
import {
  assert,
  LIMITS,
  key,
  OpenSheetError,
  uid,
  type CellRecord,
  type CellInput,
  type CellStyle,
  type CellValue,
  type Command,
  type CommandEnvelope,
  type Commit,
  type Rect,
  type Selection,
  type SheetSnapshot,
  type WorkbookSnapshot,
} from './types.js';

function notifyObserver<T extends unknown[]>(observer: (...args: T) => unknown, ...args: T) {
  try {
    const result = observer(...args);
    if (result && typeof (result as PromiseLike<unknown>).then === 'function')
      void Promise.resolve(result).catch((error) => console.error(error));
  } catch (error) {
    console.error(error);
  }
}
export interface OperationTiming {
  method: string;
  queueMs: number;
  executionMs: number;
  roundTripMs: number;
  activeAtArrival: string;
  queuedAhead: number;
  read?: ReadTiming;
}
export interface WorkbookOptions {
  database?: string;
  budgetBytes?: number;
  preserveOrphans?: boolean;
  cacheBytes?: number;
  viewportCacheBytes?: number;
  workerFactory?: () => Worker;
  operationTimeoutMs?: number;
  onProgress?: (progress: Progress) => void;
  onOperationTiming?: (timing: OperationTiming) => void;
}
export type RangeData = {
  sheetId: string;
  range: Rect;
  cells: Record<string, CellRecord>;
  calculated: Record<string, CellValue>;
  display: Record<string, string>;
  revision: number;
};
export type CellField = 'input' | 'value' | 'display' | 'style' | 'numberFormat' | 'note' | 'link';
export type CellData = {
  row: number;
  column: number;
  input?: CellInput;
  value?: CellValue;
  display?: string;
  style?: CellStyle;
  numberFormat?: string;
  note?: string;
  link?: CellRecord['link'];
};
export type CellCursor = { revision: number; sheetId: string; address: string; offset: number };
type CommandAxisState = { rows: number; columns: number; nextRowId: number; nextColumnId: number };
export type PersistenceState = {
  state: 'saving' | 'saved' | 'failed' | 'read-only';
  pending: number;
  acknowledgeable: boolean;
  error?: { code?: string; message: string };
};
export class Workbook {
  private worker: Worker;
  private initial: Promise<this>;
  private view!: EngineView;
  private pending = new Map<
    number,
    {
      resolve(value: any): void;
      reject(error: unknown): void;
      timer: ReturnType<typeof setTimeout>;
      method: string;
      sentAt?: number;
    }
  >();
  private nextId = 0;
  private disposed = false;
  private cache: ByteCache<RangeData>;
  private loading = new Map<string, Promise<void>>();
  private listeners = new Set<(commit: Commit) => void>();
  private progressListeners = new Set<(progress: Progress) => void>();
  private dataListeners = new Set<() => void>();
  private disposeListeners = new Set<() => void>();
  private persistenceListeners = new Set<(state: PersistenceState) => void>();
  private pendingWrites = 0;
  private writeFailure?: unknown;
  private batch?: Command[];
  private stagedSheets?: Map<string, SheetSnapshot>;
  private stagedAxes?: Map<string, CommandAxisState>;
  private stagedNames?: Map<string, string>;
  private stagedRemoved?: Set<string>;
  private transactionRunning = false;
  private importRunning = false;
  private generateRunning = false;
  private poisoned = false;
  private expired = false;
  private transactionCancelled = false;
  private readOnly = false;
  private writer = false;
  private channel?: BroadcastChannel;
  private remoteRefreshTimer?: ReturnType<typeof setTimeout>;
  private remoteRefreshing = false;
  private remoteRefreshAgain = false;
  private deleting = false;
  private readonly refreshOnVisible = () => {
    if (typeof document === 'undefined' || document.visibilityState === 'visible')
      this.refreshRemote();
  };
  private seen: ByteCache<{ signature: string; commit?: Commit }>;
  readonly id: string;
  constructor(
    snapshot?: WorkbookFile,
    private options: WorkbookOptions = {},
    id = snapshot?.workbookId ?? uid('wb'),
  ) {
    assert(
      options.cacheBytes === undefined ||
        (Number.isSafeInteger(options.cacheBytes) && options.cacheBytes >= 0),
      'INVALID_ARGUMENT',
      'Invalid worker cache budget',
    );
    assert(
      options.preserveOrphans === undefined || typeof options.preserveOrphans === 'boolean',
      'INVALID_ARGUMENT',
      'Invalid orphan-preservation option',
    );
    assert(
      options.viewportCacheBytes === undefined ||
        (Number.isSafeInteger(options.viewportCacheBytes) && options.viewportCacheBytes >= 0),
      'INVALID_ARGUMENT',
      'Invalid viewport cache budget',
    );
    let workerAvailable = false;
    if (!options.workerFactory) {
      try {
        workerAvailable = typeof Worker === 'function';
      } catch {
        // Access itself can be denied in restricted browser contexts.
      }
    }
    assert(
      options.workerFactory || workerAvailable,
      'WORKER_UNAVAILABLE',
      'A browser Worker or injected workerFactory is required',
    );
    let Channel: typeof BroadcastChannel | undefined;
    try {
      const candidate = globalThis.BroadcastChannel;
      if (typeof candidate === 'function') Channel = candidate;
    } catch {
      // Access itself can be denied in restricted browser contexts.
    }
    assert(
      Channel,
      'BROADCAST_CHANNEL_UNAVAILABLE',
      'BroadcastChannel is required for workbook revision coordination; use a supported browser or inject a compatible runtime',
    );
    if (!options.workerFactory && typeof window !== 'undefined') {
      let storageAvailable = false;
      try {
        storageAvailable = typeof indexedDB !== 'undefined';
      } catch {
        // Access itself can be denied in restricted browser contexts.
      }
      assert(
        storageAvailable,
        'STORAGE_UNAVAILABLE',
        'IndexedDB is required for local persistence; enable site storage or use a supported browser',
      );
      let locksAvailable = false;
      try {
        locksAvailable =
          typeof navigator !== 'undefined' && typeof navigator.locks?.request === 'function';
      } catch {
        // A restricted context can deny access to the Web Locks property itself.
      }
      assert(
        locksAvailable,
        'WEB_LOCKS_UNAVAILABLE',
        'Web Locks is required for safe multi-tab editing; open this app over HTTPS or localhost in a supported browser',
      );
    }
    this.id = id;
    const cacheBudget = options.viewportCacheBytes ?? 8 * 1024 * 1024;
    const retryBytes = Math.min(512 * 1024, Math.floor(cacheBudget / 16));
    this.seen = new ByteCache(retryBytes);
    this.cache = new ByteCache(cacheBudget - retryBytes);
    const injectedWorker = options.workerFactory?.();
    if (injectedWorker != null) {
      this.worker = injectedWorker;
    } else {
      try {
        this.worker = new Worker(new URL('./engine.worker.js', import.meta.url), {
          type: 'module',
        });
      } catch (error) {
        if (
          error instanceof Error &&
          (error.name === 'SecurityError' || error.name === 'NotAllowedError')
        )
          throw new OpenSheetError(
            'WORKER_UNAVAILABLE',
            'Worker creation was denied; check the site security policy and allow module Workers, then reopen the workbook',
          );
        throw error;
      }
    }
    this.worker.onmessage = (event) => {
      const message = event.data;
      if (message.storage) {
        if (this.view) {
          this.view.bytes = message.storage.bytes;
          this.view.cacheBytes = message.storage.cacheBytes;
        }
        return;
      }
      if (message.progress) {
        const task = this.pending.get(message.requestId);
        if (task) {
          clearTimeout(task.timer);
          task.timer = setTimeout(() => {
            this.pending.delete(message.requestId);
            task.reject(
              new OpenSheetError('WORKER_TIMEOUT', 'Worker stopped reporting import progress'),
            );
            this.worker.terminate();
            this.fail(
              new OpenSheetError(
                'WORKER_FAILED',
                'Worker stopped responding; reopen the last committed workbook',
              ),
            );
          }, this.options.operationTimeoutMs ?? 60000);
        }
        if (options.onProgress) notifyObserver(options.onProgress.bind(options), message.progress);
        for (const listener of this.progressListeners) notifyObserver(listener, message.progress);
        return;
      }
      const task = this.pending.get(message.id);
      if (!task) return;
      this.pending.delete(message.id);
      clearTimeout(task.timer);
      if (task.sentAt !== undefined && message.timing && options.onOperationTiming) {
        notifyObserver(options.onOperationTiming.bind(options), {
          method: task.method,
          queueMs: message.timing.queueMs,
          executionMs: message.timing.executionMs,
          roundTripMs: performance.now() - task.sentAt,
          activeAtArrival: message.timing.activeAtArrival,
          queuedAhead: message.timing.queuedAhead,
          ...(message.timing.read ? { read: message.timing.read } : {}),
        });
      }
      if (message.error) {
        const error = new OpenSheetError(message.error.code, message.error.message);
        task.reject(error);
        if (error.code === 'WORKER_FAILED') {
          this.worker.terminate();
          this.fail(error);
        }
      } else task.resolve(message.result);
    };
    this.worker.onerror = (event) => {
      this.worker.terminate();
      const message = this.view
        ? event.message || 'Storage worker failed'
        : `Storage worker failed to start; check the Worker asset path and site security policy, then reopen the workbook${event.message ? `: ${event.message}` : ''}`;
      this.fail(new OpenSheetError('WORKER_FAILED', message));
    };
    this.worker.onmessageerror = () => {
      this.worker.terminate();
      this.fail(new OpenSheetError('WORKER_FAILED', 'Invalid worker response'));
    };
    this.initial = this.rpc('open', [
      id,
      snapshot,
      {
        database: options.database,
        budgetBytes: options.budgetBytes,
        preserveOrphans: options.preserveOrphans,
        operationTiming: !!options.onOperationTiming,
        cacheBytes: options.cacheBytes,
      },
    ])
      .then(async ({ view, writer }) => {
        if (this.disposed)
          throw new OpenSheetError('DISPOSED', 'Workbook closed during initialization');
        this.adopt(view);
        this.writer = writer;
        try {
          this.channel = new Channel(`opensheet:${options.database ?? 'opensheet'}:${id}`);
        } catch (error) {
          if (
            error instanceof Error &&
            (error.name === 'SecurityError' || error.name === 'NotAllowedError')
          )
            throw new OpenSheetError(
              'BROADCAST_CHANNEL_UNAVAILABLE',
              'BroadcastChannel access was denied; enable site storage or use a supported browser and reopen the workbook',
            );
          throw error;
        }
        this.channel.onmessage = (event) => {
          if (event.data?.deleted === true) void this.checkRemoteDeletion();
          else this.refreshRemote();
        };
        if (typeof window !== 'undefined') window.addEventListener('focus', this.refreshOnVisible);
        if (typeof document !== 'undefined')
          document.addEventListener('visibilitychange', this.refreshOnVisible);
        // Subscribe before checking again, so a commit during initial open is not missed.
        if (!writer) {
          const latest = await this.rpc('reloadIfChanged');
          if (latest && latest.revision >= this.revision) this.adopt(latest);
        }
        return this;
      })
      .catch((error) => {
        this.dispose();
        throw error;
      });
  }
  private async checkRemoteDeletion() {
    if (this.disposed || this.deleting || this.writer || this.failure) return;
    try {
      const exists = await this.rpc('headExists');
      if (this.disposed || this.writer || this.failure) return;
      if (exists) {
        this.refreshRemote();
        return;
      }
      this.failRemote(
        new OpenSheetError('WORKBOOK_DELETED', 'This workbook was deleted in another tab'),
      );
    } catch (error) {
      if (!this.disposed && !this.writer && !this.failure) this.failRemote(error);
    }
  }
  private failRemote(error: unknown) {
    this.cache.clear();
    this.fail(error);
    for (const listener of this.dataListeners) notifyObserver(listener);
  }
  private refreshRemote() {
    if (this.disposed || this.deleting || this.writer || this.failure) return;
    if (this.remoteRefreshing || this.remoteRefreshTimer) {
      this.remoteRefreshAgain = true;
      return;
    }
    this.remoteRefreshing = true;
    void this.rpc('reloadIfChanged')
      .then((view) => {
        if (this.disposed || this.writer || this.failure) return;
        if (!view) return;
        const previousRevision = this.revision;
        if (view.revision <= previousRevision) return;
        this.adopt(view);
        this.emit({
          commandId: uid('remote'),
          previousRevision,
          revision: this.revision,
          source: 'remote',
          commands: [],
          changedRanges: [],
        });
      })
      .catch((error) => {
        if (this.disposed || this.writer || this.failure) return;
        if (error?.code === 'REVISION_CONFLICT') {
          if (!this.remoteRefreshTimer)
            this.remoteRefreshTimer = setTimeout(() => {
              this.remoteRefreshTimer = undefined;
              this.remoteRefreshAgain = false;
              this.refreshRemote();
            }, 100);
        } else {
          this.remoteRefreshAgain = false;
          this.failRemote(error);
        }
      })
      .finally(() => {
        this.remoteRefreshing = false;
        if (this.remoteRefreshAgain && !this.remoteRefreshTimer) {
          this.remoteRefreshAgain = false;
          this.refreshRemote();
        }
      });
  }
  private failure?: unknown;
  private fail(error: unknown) {
    this.failure = error;
    for (const p of this.pending.values()) {
      clearTimeout(p.timer);
      p.reject(error);
    }
    this.pending.clear();
    this.emitPersistence();
  }
  private rpc(
    method: string,
    args: unknown[] = [],
    transfer: Transferable[] = [],
    signal?: AbortSignal,
  ): Promise<any> {
    assert(!this.disposed, 'DISPOSED', 'Workbook is disposed');
    assert(!this.deleting || method === 'clear', 'BUSY', 'Workbook deletion is in progress');
    if (this.failure) return Promise.reject(this.failure);
    if (signal?.aborted)
      return Promise.reject(new OpenSheetError('ABORTED', 'Operation cancelled'));
    const id = ++this.nextId;
    // Bind metadata elision to what this caller actually held when it sent the request.
    const knownView = this.view;
    const cancel = () => this.worker.postMessage({ id: 0, method: 'cancelRequest', args: [id] });
    const write = ['execute', 'history', 'generate', 'importJSON', 'importBinary'].includes(method);
    if (write) {
      this.pendingWrites++;
    }
    const request = new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(
          new OpenSheetError('WORKER_TIMEOUT', `Worker request ${method} exceeded its time budget`),
        );
        this.worker.terminate();
        this.fail(
          new OpenSheetError(
            'WORKER_FAILED',
            'Storage worker stopped responding; reopen the last committed workbook',
          ),
        );
      }, this.options.operationTimeoutMs ?? 60000);
      this.pending.set(id, {
        resolve,
        reject,
        timer,
        method,
        sentAt: this.options.onOperationTiming ? performance.now() : undefined,
      });
      try {
        this.worker.postMessage({ id, method, args, metadataId: knownView?.metadataId }, transfer);
        signal?.addEventListener('abort', cancel, { once: true });
        if (signal?.aborted) cancel();
      } catch (error) {
        this.pending.delete(id);
        clearTimeout(timer);
        reject(error);
      }
    });
    // Submit before notifying synchronous observers, which may issue another write.
    if (write) this.emitPersistence();
    const result = (
      signal
        ? request.then(
            (value) => {
              signal.removeEventListener('abort', cancel);
              return value;
            },
            (error) => {
              signal.removeEventListener('abort', cancel);
              throw error;
            },
          )
        : request
    ).then((value) => {
      if (knownView && value && typeof value === 'object') {
        const container = value as any;
        const view = 'view' in container ? container.view : container;
        if (view && view.metadataId === knownView.metadataId && !view.file && !view.snapshot)
          view.snapshot = knownView.snapshot;
      }
      return value;
    });
    if (!write) return result;
    return result.then(
      (result) => {
        // The caller adopts the acknowledged view before clearing the pending write.
        return result;
      },
      (error) => {
        this.pendingWrites--;
        this.writeFailure = error;
        this.emitPersistence();
        throw error;
      },
    );
  }
  ready() {
    return this.initial;
  }
  get revision() {
    return this.view.snapshot.revision;
  }
  get metadata() {
    assert(
      !this.disposed && !this.expired,
      'DISPOSED',
      'Workbook or transaction scope is disposed',
    );
    return this.view.snapshot;
  }
  get isReadOnly() {
    return this.readOnly || !this.writer || this.metadata.extensions['opensheet.readOnly'] === true;
  }
  get canRequestWriteLock() {
    return (
      !this.failure &&
      !this.writer &&
      !this.readOnly &&
      this.metadata.extensions['opensheet.readOnly'] !== true
    );
  }
  async tryAcquireWriteLock(): Promise<boolean> {
    assert(!this.disposed, 'DISPOSED', 'Workbook is disposed');
    if (this.failure) throw this.failure;
    assert(
      !this.readOnly && this.metadata.extensions['opensheet.readOnly'] !== true,
      'READ_ONLY',
      'Workbook is explicitly read only',
    );
    if (this.writer) return true;
    const result = await this.rpc('tryAcquireWriteLock');
    if (!result.writer) return false;
    const previousRevision = this.revision;
    try {
      this.adopt(result.view);
    } catch (error) {
      this.worker.terminate();
      this.fail(
        new OpenSheetError('WORKER_FAILED', 'Could not adopt the recovered workbook; reopen it'),
      );
      throw error;
    }
    this.writer = true;
    if (previousRevision !== this.revision) {
      this.emit({
        commandId: uid('remote'),
        previousRevision,
        revision: this.revision,
        source: 'remote',
        commands: [],
        changedRanges: [],
      });
    }
    this.emitPersistence();
    return true;
  }
  get canUndo() {
    return this.view.canUndo;
  }
  get canRedo() {
    return this.view.canRedo;
  }
  async storageStats(): Promise<{
    bytes: number;
    cacheBytes: number;
    viewportCacheBytes: number;
    budgetBytes: number;
  }> {
    return { ...(await this.rpc('stats')), viewportCacheBytes: this.cache.bytes + this.seen.bytes };
  }
  get stats() {
    return {
      bytes: this.view.bytes,
      cacheBytes: this.view.cacheBytes,
      viewportCacheBytes: this.cache.bytes + this.seen.bytes,
      budgetBytes: this.view.budget,
    };
  }
  setReadOnly(value: boolean) {
    if (this.readOnly === value) return;
    this.readOnly = value;
    this.emitPersistence();
  }
  get persistenceState(): PersistenceState {
    const failure = this.failure ?? this.writeFailure;
    const error =
      failure === undefined
        ? undefined
        : {
            ...((failure as { code?: unknown })?.code &&
            typeof (failure as { code?: unknown }).code === 'string'
              ? { code: (failure as { code: string }).code }
              : {}),
            message: failure instanceof Error ? failure.message : String(failure),
          };
    return {
      state: error
        ? 'failed'
        : this.pendingWrites
          ? 'saving'
          : this.isReadOnly
            ? 'read-only'
            : 'saved',
      pending: this.pendingWrites,
      acknowledgeable: !this.failure && this.writeFailure !== undefined,
      ...(error ? { error } : {}),
    };
  }
  onPersistence(fn: (state: PersistenceState) => void) {
    this.persistenceListeners.add(fn);
    return () => {
      this.persistenceListeners.delete(fn);
    };
  }
  acknowledgePersistenceFailure() {
    assert(!this.disposed, 'DISPOSED', 'Workbook is disposed');
    this.writeFailure = undefined;
    this.emitPersistence();
  }
  private emitPersistence() {
    if (this.disposed || !this.view || !this.persistenceListeners.size) return;
    for (const listener of this.persistenceListeners) {
      try {
        notifyObserver(listener, this.persistenceState);
      } catch (error) {
        console.error(error);
      }
    }
  }
  getSheets() {
    const ids = this.metadata.sheetOrder;
    if (!this.stagedSheets?.size && !this.stagedRemoved?.size)
      return ids.map((id) => new Sheet(this, id));
    const retained = ids.filter((id) => !this.stagedRemoved?.has(id));
    const retainedIds = new Set(retained);
    const staged = [...(this.stagedSheets?.keys() ?? [])].filter((id) => !retainedIds.has(id));
    return [...retained, ...staged].map((id) => new Sheet(this, id));
  }
  getSheetById(id: string) {
    const committed = this.metadata.sheets.some((s) => s.id === id);
    return this.stagedSheets?.has(id) || (committed && !this.stagedRemoved?.has(id))
      ? new Sheet(this, id)
      : undefined;
  }
  getSheetByName(name: string) {
    return this.getSheets().find((s) => s.name.toLowerCase() === name.toLowerCase());
  }
  sheetData(id: string) {
    const sh = this.metadata.sheets.find((s) => s.id === id) ?? this.stagedSheets?.get(id);
    assert(sh, 'INVALID_ARGUMENT', 'Unknown sheet');
    return sh;
  }
  /** @internal Names used for commands and lookups inside a transaction. */
  sheetCommandName(id: string): string {
    const sheet = this.sheetData(id);
    return this.stagedNames?.get(id) ?? sheet.name;
  }
  /** @internal Dimensions and ID counters for commands; reads use committed sheetData. */
  sheetCommandState(id: string): CommandAxisState {
    const committed = this.sheetData(id);
    const source = this.stagedSheets?.get(id) ?? committed;
    return (
      this.stagedAxes?.get(id) ?? {
        rows: source.rowOrder.length,
        columns: source.columnOrder.length,
        nextRowId: source.nextRowId,
        nextColumnId: source.nextColumnId,
      }
    );
  }
  /** @internal Whether commands are being collected by a transaction scope. */
  get isTransactionScope() {
    return !!this.batch;
  }
  getStyle(id?: string) {
    return id ? { ...this.metadata.styles[id] } : {};
  }
  usedRange(id: string) {
    return this.view.used[id];
  }
  filteredRows(id: string) {
    return this.view.filtered[id] ?? [];
  }
  isLoaded(sheet: string, row: number, column: number) {
    return !!this.cache.get(blockAddress(sheet, row, column));
  }
  peekCell(sheet: string, row: number, column: number) {
    const sh = this.sheetData(sheet);
    return this.cache.get(blockAddress(sheet, row, column))?.cells[
      key(sh.rowOrder[row], sh.columnOrder[column])
    ];
  }
  peekValue(sheet: string, row: number, column: number): CellValue {
    const sh = this.sheetData(sheet);
    const data = this.cache.get(blockAddress(sheet, row, column));
    return data
      ? (data.calculated[key(sh.rowOrder[row], sh.columnOrder[column])] ?? null)
      : { error: '#LOADING' };
  }
  peekDisplay(sheet: string, row: number, column: number) {
    const sh = this.sheetData(sheet);
    const data = this.cache.get(blockAddress(sheet, row, column));
    return data ? (data.display[key(sh.rowOrder[row], sh.columnOrder[column])] ?? '') : '…';
  }
  private async readCellData(sheet: string, row: number, column: number, values: boolean) {
    const revision = this.revision;
    const sh = this.sheetData(sheet);
    const address = key(sh.rowOrder[row], sh.columnOrder[column]);
    const data = await this.readRange(
      sheet,
      {
        startRow: row,
        endRow: row + 1,
        startColumn: column,
        endColumn: column + 1,
      },
      values,
    );
    assert(data.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during read');
    return { data, address };
  }
  async getCell(sheet: string, row: number, column: number) {
    const { data, address } = await this.readCellData(sheet, row, column, false);
    return data.cells[address];
  }
  async value(sheet: string, row: number, column: number) {
    const { data, address } = await this.readCellData(sheet, row, column, true);
    return data.calculated[address] ?? null;
  }
  async display(sheet: string, row: number, column: number) {
    const { data, address } = await this.readCellData(sheet, row, column, true);
    return data.display[address] ?? '';
  }
  async prefetch(sheet: string, range: Rect) {
    const sh = this.sheetData(sheet);
    checkBounds(sh, range);
    const revision = this.revision;
    const tasks: Promise<void>[] = [];
    for (
      let row = Math.floor(range.startRow / BLOCK_ROWS) * BLOCK_ROWS;
      row < range.endRow;
      row += BLOCK_ROWS
    )
      for (
        let col = Math.floor(range.startColumn / BLOCK_COLUMNS) * BLOCK_COLUMNS;
        col < range.endColumn;
        col += BLOCK_COLUMNS
      ) {
        const a = blockAddress(sheet, row, col);
        if (this.cache.get(a)) continue;
        let task = this.loading.get(a);
        if (!task) {
          task = this.rpc('read', [
            sheet,
            {
              startRow: row,
              endRow: Math.min(row + BLOCK_ROWS, sh.rowOrder.length),
              startColumn: col,
              endColumn: Math.min(col + BLOCK_COLUMNS, sh.columnOrder.length),
            },
            true,
          ])
            .then((data: RangeData) => {
              if (
                !this.disposed &&
                !this.failure &&
                revision === this.revision &&
                data.revision === revision
              ) {
                // Renderer-facing peeks share these objects; freeze once at the
                // cache boundary so callers cannot mutate cached workbook state.
                this.cache.set(a, freeze(data, true));
                for (const fn of this.dataListeners) notifyObserver(fn);
              }
            })
            .finally(() => {
              if (this.loading.get(a) === task) this.loading.delete(a);
            });
          this.loading.set(a, task);
        }
        tasks.push(task);
      }
    await Promise.all(tasks);
  }
  async readRange(sheet: string, range: Rect, values = true): Promise<RangeData> {
    checkBounds(this.sheetData(sheet), range);
    const data: RangeData = await this.rpc('read', [sheet, range, values]);
    assert(data.revision === this.revision, 'REVISION_CONFLICT', 'Workbook changed during read');
    return data;
  }
  async readCells(sheetId: string, range: Rect, options: { fields?: CellField[] } = {}) {
    checkBounds(this.sheetData(sheetId), range);
    const fields = options.fields ?? ['input', 'value'];
    assert(
      fields.every((field) =>
        ['input', 'value', 'display', 'style', 'numberFormat', 'note', 'link'].includes(field),
      ),
      'INVALID_ARGUMENT',
      'Unknown cell field',
    );
    const data: { revision: number; sheetId: string; range: Rect; cells: CellData[] } =
      await this.rpc('readCells', [sheetId, range, fields]);
    assert(data.revision === this.revision, 'REVISION_CONFLICT', 'Workbook changed during read');
    return data;
  }
  async scanCells(
    sheetId: string,
    options: { cursor?: CellCursor; limit?: number; fields?: CellField[] } = {},
  ) {
    const fields = options.fields ?? ['input'];
    assert(
      fields.every((field) =>
        ['input', 'value', 'display', 'style', 'numberFormat', 'note', 'link'].includes(field),
      ),
      'INVALID_ARGUMENT',
      'Unknown cell field',
    );
    const result: { revision: number; cells: CellData[]; cursor?: CellCursor } = await this.rpc(
      'scanCells',
      [sheetId, { ...options, fields }],
    );
    assert(result.revision === this.revision, 'REVISION_CONFLICT', 'Workbook changed during read');
    return result;
  }
  async *streamRange(
    sheet: string,
    range: Rect,
    options: { signal?: AbortSignal; values?: boolean } = {},
  ) {
    const bounds = { ...range };
    const signal = options.signal;
    const values = options.values ?? true;
    checkBounds(this.sheetData(sheet), bounds);
    const revision = this.revision;
    const width = bounds.endColumn - bounds.startColumn,
      rows = Math.max(1, Math.floor(16384 / width));
    for (let row = bounds.startRow; row < bounds.endRow; row += rows) {
      signal?.throwIfAborted();
      assert(
        revision === this.revision,
        'REVISION_CONFLICT',
        'Workbook changed during streaming read',
      );
      const data = await this.readRange(
        sheet,
        { ...bounds, startRow: row, endRow: Math.min(row + rows, bounds.endRow) },
        values,
      );
      assert(
        data.revision === revision && this.revision === revision,
        'REVISION_CONFLICT',
        'Workbook changed during streaming read',
      );
      yield data;
    }
  }
  private adopt(view: EngineView) {
    const wire = view as EngineView & { file?: WorkbookFile; styleIds?: string[] };
    assert(
      wire.file || wire.snapshot || this.view?.metadataId === view.metadataId,
      'CORRUPT_STORAGE',
      'Worker response omitted required workbook metadata',
    );
    const snapshot = wire.file
      ? decodeFileMetadata(wire.file)
      : (wire.snapshot ?? this.view?.snapshot);
    assert(snapshot, 'CORRUPT_STORAGE', 'Missing workbook metadata');
    if (wire.file && wire.styleIds)
      snapshot.styles = Object.fromEntries(
        wire.styleIds.map((id, i) => [id, wire.file!.styles[i]]),
      );
    // These objects are returned directly by synchronous metadata accessors.
    // Protect them once per adopted view rather than copying large row lists on
    // every renderer layout.
    for (const bounds of Object.values(view.used)) Object.freeze(bounds);
    for (const rows of Object.values(view.filtered)) Object.freeze(rows);
    this.view = { ...view, snapshot: freeze({ ...snapshot, revision: view.revision }, true) };
    this.cache.clear();
    this.loading.clear();
  }
  onProgress(fn: (progress: Progress) => void) {
    this.progressListeners.add(fn);
    return () => {
      this.progressListeners.delete(fn);
    };
  }
  onCommit(fn: (commit: Commit) => void) {
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }
  onData(fn: () => void) {
    this.dataListeners.add(fn);
    return () => {
      this.dataListeners.delete(fn);
    };
  }
  onDispose(fn: () => void) {
    assert(
      !this.disposed && !this.expired,
      'DISPOSED',
      'Workbook or transaction scope is disposed',
    );
    this.disposeListeners.add(fn);
    return () => {
      this.disposeListeners.delete(fn);
    };
  }
  private emit(commit: Commit) {
    for (const fn of this.listeners) {
      try {
        notifyObserver(fn, structuredClone(commit));
      } catch (error) {
        console.error(error);
      }
    }
  }
  private notifyRevision() {
    try {
      this.channel?.postMessage({ revision: this.revision });
    } catch {
      // The head has already committed. Preserve its acknowledgement, but stop
      // further work because other tabs may not learn about this revision.
      this.worker.terminate();
      this.fail(
        new OpenSheetError(
          'SYNC_UNAVAILABLE',
          'Tab synchronization failed. Reopen the workbook and check pending changes.',
        ),
      );
    }
  }
  private accept(result: { view: EngineView; commit?: Commit; duplicate?: boolean }) {
    if (result.duplicate) return result.commit;
    // A delayed no-op response cannot replace a newer revision already adopted.
    if (!result.commit && result.view.revision < this.revision) return;
    if (
      !result.commit &&
      result.view.revision === this.revision &&
      result.view.metadataId === this.view.metadataId
    ) {
      // A no-op at the already adopted revision cannot change workbook data.
      // Refresh accounting without discarding loaded or in-flight viewport pages.
      this.view = {
        ...result.view,
        snapshot: this.view.snapshot,
        used: this.view.used,
        filtered: this.view.filtered,
      };
      return;
    }
    // A successful acknowledgement may also arrive after a newer response.
    // Deliver its commit to observers, but never regress the adopted view.
    if (result.view.revision >= this.revision) this.adopt(result.view);
    if (result.commit) {
      this.emit(result.commit);
      this.notifyRevision();
    }
    return result.commit;
  }
  async execute(input: Command | CommandEnvelope): Promise<Commit | undefined> {
    assert(!this.importRunning && !this.generateRunning, 'BUSY', 'Workbook replacement is running');
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    if (this.batch) {
      try {
        assert(!('protocolVersion' in input), 'INVALID_ARGUMENT', 'Envelopes cannot be batched');
        const command = validateCommand(input);
        this.batch.push(command);
        if (command.type === 'core.sheet.add') {
          const { sheet } = command.payload as { sheet: SheetSnapshot };
          this.stagedSheets?.set(sheet.id, freeze(sheet, true));
          this.stagedNames?.set(sheet.id, sheet.name);
        } else if (command.type === 'core.sheet.remove') {
          const { sheetId } = command.payload as { sheetId: string };
          this.stagedSheets?.delete(sheetId);
          this.stagedAxes?.delete(sheetId);
          this.stagedNames?.delete(sheetId);
          if (this.metadata.sheets.some((sheet) => sheet.id === sheetId))
            this.stagedRemoved?.add(sheetId);
        } else if (command.type === 'core.sheet.rename') {
          const { sheetId, name } = command.payload as { sheetId: string; name: string };
          this.stagedNames?.set(sheetId, name);
          const sheet = this.stagedSheets?.get(sheetId);
          if (sheet) {
            this.stagedSheets?.set(sheetId, freeze({ ...sheet, name }, true));
          }
        } else if (command.type === 'core.axis.insert' || command.type === 'core.axis.delete') {
          const p = command.payload as {
            sheetId: string;
            axis: 'row' | 'column';
            index: number;
            ids?: number[];
            count?: number;
          };
          const sheet = this.sheetCommandState(p.sheetId);
          const row = p.axis === 'row';
          const length = row ? sheet.rows : sheet.columns;
          const insert = command.type === 'core.axis.insert';
          const count = insert ? p.ids!.length : p.count!;
          assert(
            p.index <= (insert ? length : length - count) && count > 0,
            'INVALID_RANGE',
            'Invalid structural range',
          );
          assert(
            insert ? length + count <= (row ? LIMITS.rows : LIMITS.columns) : length - count >= 1,
            'LIMIT_EXCEEDED',
            'Invalid sheet size',
          );
          const next = row ? sheet.nextRowId : sheet.nextColumnId;
          if (insert)
            assert(
              p.ids!.every((id, i) => id === next + i) && next + count <= 0x100000000,
              'INVALID_ARGUMENT',
              'Axis IDs must be allocated monotonically',
            );
          const staged = {
            ...sheet,
            [row ? 'rows' : 'columns']: length + (insert ? count : -count),
            ...(insert ? { [row ? 'nextRowId' : 'nextColumnId']: next + count } : {}),
          };
          Object.freeze(staged);
          this.stagedAxes?.set(p.sheetId, staged);
        }
        return;
      } catch (error) {
        this.poisoned = true;
        throw error;
      }
    }
    assert(!this.transactionRunning, 'BUSY', 'Use the scoped workbook inside a transaction');
    if (input.type === 'core.cells.copy' || input.type === 'core.cells.fillDown')
      assert(this.pendingWrites === 0, 'BUSY', 'Wait for pending writes before copying or filling');
    let signature = '';
    const commandId = 'protocolVersion' in input ? input.commandId : undefined;
    if ('protocolVersion' in input) {
      assert(
        input.protocolVersion === 1 &&
          input.workbookId === this.id &&
          input.commandId.length > 0 &&
          input.commandId.length <= 100,
        'INVALID_ARGUMENT',
        'Invalid command envelope',
      );
      signature = JSON.stringify(input);
      const seen = this.seen.get(input.commandId);
      if (seen) {
        assert(seen.signature === signature, 'INVALID_ARGUMENT', 'Command id reused');
        return seen.commit ? structuredClone(seen.commit) : undefined;
      }
      assert(input.baseRevision === this.revision, 'REVISION_CONFLICT', 'Workbook changed');
    }
    const result = await this.rpc('execute', [[input], 'edit']);
    try {
      const commit = this.accept(result);
      if (commandId !== undefined) {
        const bytes = encodedBytes({ signature, commit }) * 4;
        if (bytes <= this.seen.budget)
          this.seen.set(
            commandId,
            { signature, commit: commit ? structuredClone(commit) : undefined },
            bytes,
          );
      }
      return commit;
    } catch (error) {
      this.writeFailure = error;
      throw error;
    } finally {
      this.pendingWrites--;
      this.emitPersistence();
    }
  }
  async transaction<T>(
    options: { label?: string },
    fn: (workbook: Workbook) => T | Promise<T>,
  ): Promise<T> {
    assert(
      !this.batch && !this.transactionRunning && !this.isReadOnly,
      'INVALID_ARGUMENT',
      'Transactions cannot nest and require write access',
    );
    assert(
      !this.importRunning && !this.generateRunning && this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    this.transactionRunning = true;
    this.transactionCancelled = false;
    this.pendingWrites++;
    this.emitPersistence();
    const scoped = Object.create(this) as Workbook;
    scoped.batch = [];
    scoped.stagedSheets = new Map();
    scoped.stagedAxes = new Map();
    scoped.stagedNames = new Map();
    scoped.stagedRemoved = new Set();
    scoped.poisoned = false;
    scoped.transactionRunning = false;
    const scopeStops = new Set<() => void>();
    const subscribe = <T extends unknown[]>(
      add: (listener: (...args: T) => void) => () => void,
      listener: (...args: T) => void,
    ) => {
      assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
      const remove = add((...args) => {
        if (!scoped.expired) listener(...args);
      });
      let active = true;
      const stop = () => {
        if (!active) return;
        active = false;
        scopeStops.delete(stop);
        remove();
      };
      scopeStops.add(stop);
      return stop;
    };
    scoped.onCommit = (listener) => subscribe((fn) => this.onCommit(fn), listener);
    scoped.onData = (listener) => subscribe((fn) => this.onData(fn), listener);
    scoped.onDispose = (listener) => subscribe((fn) => this.onDispose(fn), listener);
    scoped.onProgress = (listener) => subscribe((fn) => this.onProgress(fn), listener);
    scoped.onPersistence = (listener) => subscribe((fn) => this.onPersistence(fn), listener);
    scoped.rpc = async (...args) => {
      assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
      const result = await this.rpc(...args);
      assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
      return result;
    };
    scoped.ready = async () => {
      assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
      await this.ready();
      assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
      return scoped;
    };
    scoped.calculation = {
      calculate: async () => {
        // Calculation reads the committed revision; newly staged sheets do not
        // exist in that revision and must not be included in this scan.
        for (const sh of this.getSheets())
          for await (const _data of scoped.streamRange(sh.id, scoped.usedRange(sh.id))) {
          }
        assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
        return { revision: scoped.revision };
      },
      getState: () => {
        assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
        return { revision: scoped.revision, state: 'ready' as const };
      },
    };
    scoped.cancel = () => {
      assert(!scoped.expired, 'DISPOSED', 'Transaction scope is disposed');
      this.cancel();
    };
    scoped.dispose = () => {
      throw new OpenSheetError('INVALID_ARGUMENT', 'Cannot dispose a transaction scope');
    };
    try {
      const value = await fn(scoped);
      assert(!scoped.poisoned, 'INVALID_ARGUMENT', 'Transaction contains a failed command');
      assert(!this.transactionCancelled, 'ABORTED', 'Transaction cancelled');
      if (scoped.batch.length) {
        const result = await this.rpc('execute', [scoped.batch, options.label ?? 'transaction']);
        try {
          this.accept(result);
        } finally {
          this.pendingWrites--;
        }
      }
      return value;
    } catch (error) {
      this.writeFailure = error;
      this.emitPersistence();
      throw error;
    } finally {
      scoped.batch = undefined;
      scoped.stagedSheets?.clear();
      scoped.stagedAxes?.clear();
      scoped.stagedNames?.clear();
      scoped.stagedRemoved?.clear();
      scoped.stagedSheets = undefined;
      scoped.stagedAxes = undefined;
      scoped.stagedNames = undefined;
      scoped.stagedRemoved = undefined;
      scoped.expired = true;
      for (const stop of scopeStops) stop();
      this.transactionRunning = false;
      this.pendingWrites--;
      this.emitPersistence();
    }
  }
  async undo() {
    assert(!this.batch && !this.transactionRunning, 'BUSY', 'Cannot undo during a transaction');
    assert(
      !this.importRunning && !this.generateRunning && this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    const result = await this.rpc('history', [false]);
    try {
      return !!this.accept(result);
    } catch (error) {
      this.writeFailure = error;
      throw error;
    } finally {
      this.pendingWrites--;
      this.emitPersistence();
    }
  }
  async redo() {
    assert(!this.batch && !this.transactionRunning, 'BUSY', 'Cannot redo during a transaction');
    assert(
      !this.importRunning && !this.generateRunning && this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    const result = await this.rpc('history', [true]);
    try {
      return !!this.accept(result);
    } catch (error) {
      this.writeFailure = error;
      throw error;
    } finally {
      this.pendingWrites--;
      this.emitPersistence();
    }
  }
  async addSheet(name: string, options: { rows?: number; columns?: number } = {}) {
    const sh = createSheet(name, options.rows, options.columns);
    await this.execute({ type: 'core.sheet.add', payload: { sheet: sh } });
    return this.getSheetById(sh.id)!;
  }
  removeSheet(id: string) {
    return this.execute({ type: 'core.sheet.remove', payload: { sheetId: id } });
  }
  renameSheet(id: string, name: string) {
    return this.execute({ type: 'core.sheet.rename', payload: { sheetId: id, name } });
  }
  replaceText(sheetId: string, query: string, replacement: string) {
    return this.execute({ type: 'core.cells.replace', payload: { sheetId, query, replacement } });
  }
  /** Find a visible data-region boundary at the current revision. */
  async navigateData(
    sheet: string,
    row: number,
    column: number,
    axis: 'row' | 'column',
    direction: -1 | 1,
    options: { signal?: AbortSignal } = {},
  ) {
    assert(
      options && typeof options === 'object' && !Array.isArray(options),
      'INVALID_ARGUMENT',
      'Invalid navigation options',
    );
    const signal = options.signal;
    assert(
      signal === undefined ||
        (signal !== null &&
          typeof signal === 'object' &&
          typeof signal.aborted === 'boolean' &&
          typeof signal.addEventListener === 'function' &&
          typeof signal.removeEventListener === 'function'),
      'INVALID_ARGUMENT',
      'Invalid navigation abort signal',
    );
    validateNavigationArguments(sheet, row, column, axis, direction);
    const revision = this.revision;
    const result = (await this.rpc(
      'navigateData',
      [sheet, row, column, axis, direction],
      [],
      signal,
    )) as { revision: number; row: number; column: number };
    assert(
      result.revision === revision && this.revision === revision,
      'REVISION_CONFLICT',
      'Workbook changed during navigation',
    );
    return result;
  }
  async search(
    sheet: string,
    query: string,
    after = -1,
    options: { signal?: AbortSignal; direction?: 'next' | 'previous' } = {},
  ) {
    assert(
      options && typeof options === 'object' && !Array.isArray(options),
      'INVALID_ARGUMENT',
      'Invalid search options',
    );
    const signal = options.signal;
    assert(
      signal === undefined ||
        (signal !== null &&
          typeof signal === 'object' &&
          typeof signal.aborted === 'boolean' &&
          typeof signal.addEventListener === 'function' &&
          typeof signal.removeEventListener === 'function'),
      'INVALID_ARGUMENT',
      'Invalid search abort signal',
    );
    const direction = options.direction;
    validateSearchArguments(sheet, query, after, direction);
    return this.rpc('search', [sheet, query, after, !!signal, direction], [], signal) as Promise<{
      match?: { row: number; column: number };
      count: number;
    }>;
  }
  async generate(count: number, budget = this.view.budget) {
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    assert(
      !this.batch &&
        !this.transactionRunning &&
        !this.importRunning &&
        !this.generateRunning &&
        this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    this.generateRunning = true;
    let acknowledged = false;
    try {
      const view = await this.rpc('generate', [count, budget]);
      acknowledged = true;
      this.adopt(view);
      this.notifyRevision();
    } catch (error) {
      this.writeFailure = error;
      throw error;
    } finally {
      this.generateRunning = false;
      if (acknowledged) this.pendingWrites--;
      this.emitPersistence();
    }
  }
  getMetadata() {
    return {
      workbookId: this.id,
      revision: this.revision,
      dateSystem: this.metadata.dateSystem,
      sheets: this.getSheets().map((sh) => ({
        id: sh.id,
        name: sh.name,
        rows: sh.rowCount,
        columns: sh.columnCount,
        usedRange: this.usedRange(sh.id),
      })),
    };
  }
  getAxes(sheetId: string, axis: 'row' | 'column', start: number, count: number) {
    const sh = this.sheetData(sheetId),
      ids = axis === 'row' ? sh.rowOrder : sh.columnOrder;
    assert(
      Number.isInteger(start) &&
        Number.isInteger(count) &&
        start >= 0 &&
        count >= 0 &&
        count <= 16384 &&
        start + count <= ids.length,
      'INVALID_RANGE',
      'Invalid axis window',
    );
    return { revision: this.revision, start, ids: ids.slice(start, start + count) };
  }
  async importJSON(
    input: WorkbookFile | string | Blob | ReadableStream<Uint8Array>,
    options: { signal?: AbortSignal } = {},
  ) {
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    assert(
      !this.importRunning && !this.transactionRunning && this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    options.signal?.throwIfAborted();
    this.importRunning = true;
    const cancel = () => this.cancel();
    options.signal?.addEventListener('abort', cancel, { once: true });
    let acknowledged = false;
    try {
      const view = await this.rpc(
        'importJSON',
        [input],
        input instanceof ReadableStream ? [input] : [],
      );
      acknowledged = true;
      this.adopt(view);
      this.notifyRevision();
    } catch (error) {
      this.writeFailure = error;
      throw error;
    } finally {
      this.importRunning = false;
      options.signal?.removeEventListener('abort', cancel);
      if (acknowledged) this.pendingWrites--;
      this.emitPersistence();
    }
  }
  private async *filePages(revision: number, signal?: AbortSignal) {
    let after: string | undefined;
    while (true) {
      assert(!signal?.aborted, 'ABORTED', 'Export aborted');
      assert(this.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during export');
      const addresses: string[] = await this.rpc('exportIndex', [revision, after, 32]);
      if (!addresses.length) break;
      for (const address of addresses) {
        assert(!signal?.aborted, 'ABORTED', 'Export aborted');
        const encoded: EncodedBlock = await this.rpc('exportBlock', [address, revision]);
        const block = decodeBlock(encoded);
        yield { address, block, encoded };
      }
      after = addresses.at(-1);
    }
    assert(this.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during export');
  }
  private async exportMetadata(revision: number) {
    const ids: string[] = await this.rpc('exportStyles', [revision]);
    assert(this.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    const file = encodeWorkbookFile(this.metadata);
    file.styles = ids.map((id) => this.metadata.styles[id]);
    return { file, ids };
  }
  async toJSON(options: { signal?: AbortSignal } = {}): Promise<WorkbookFile> {
    const revision = this.revision,
      metadata = await this.exportMetadata(revision),
      file = structuredClone(metadata.file),
      ids = metadata.ids;
    const styles = new Map(ids.map((id, i) => [id, i]));
    const sheets = new Map(file.sheets.map((sh) => [sh.id, sh]));
    for await (const { address, block } of this.filePages(revision, options.signal)) {
      const [id, row, column] = address.split('/');
      sheets.get(id)!.blocks.push({
        row: +row,
        column: +column,
        data: encodeFileCells(
          Object.values(block.cells),
          styles,
          this.view.importRevision === revision,
        ),
      });
    }
    for (const sheet of file.sheets)
      sheet.blocks.sort((a, b) => a.row - b.row || a.column - b.column);
    return file;
  }
  async *streamJSON(options: { signal?: AbortSignal } = {}): AsyncGenerator<string> {
    const revision = this.revision,
      { file, ids } = await this.exportMetadata(revision);
    const styles = new Map(ids.map((id, i) => [id, i]));
    const { sheets, ...header } = file;
    const check = () => {
      assert(!options.signal?.aborted, 'ABORTED', 'Export aborted');
      assert(revision === this.revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    };
    check();
    yield JSON.stringify(header).slice(0, -1) + ',"sheets":[';
    for (let i = 0; i < sheets.length; i++) {
      check();
      const { blocks, ...meta } = sheets[i];
      yield (i ? ',' : '') + JSON.stringify(meta).slice(0, -1) + ',"blocks":[';
      let first = true,
        after: string | undefined;
      while (true) {
        check();
        const addresses: string[] = await this.rpc('exportIndex', [revision, after, 32, meta.id]);
        if (!addresses.length) break;
        for (const address of addresses) {
          check();
          const block = decodeBlock(await this.rpc('exportBlock', [address, revision]));
          check();
          const [, row, column] = address.split('/');
          yield (first ? '' : ',') +
            JSON.stringify({
              row: +row,
              column: +column,
              data: encodeFileCells(
                Object.values(block.cells),
                styles,
                this.view.importRevision === revision,
              ),
            });
          first = false;
        }
        after = addresses.at(-1);
      }
      yield ']}';
    }
    check();
    yield ']}';
  }
  async *streamBinary(options: { signal?: AbortSignal } = {}) {
    const revision = this.revision,
      { file: metadata, ids } = await this.exportMetadata(revision),
      book = this;
    const styles = new Map(ids.map((id, i) => [id, `s${i}`]));
    async function* pages() {
      for (let sheet = 0; sheet < metadata.sheets.length; sheet++) {
        const id = metadata.sheets[sheet].id;
        let after: string | undefined;
        while (true) {
          assert(!options.signal?.aborted, 'ABORTED', 'Export aborted');
          const addresses: string[] = await book.rpc('exportIndex', [revision, after, 32, id]);
          if (!addresses.length) break;
          for (const address of addresses) {
            const page = decodeBlock(await book.rpc('exportBlock', [address, revision])),
              cells = page.cells;
            for (const cell of Object.values(cells)) {
              if (cell.styleId !== undefined) cell.styleId = styles.get(cell.styleId)!;
              if (book.view.importRevision !== revision) delete cell.cached;
            }
            yield {
              sheet,
              row: page.row,
              column: page.column,
              bytes: encodeBlock(cells, [], [], {}),
            };
          }
          after = addresses.at(-1);
        }
      }
      assert(book.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    }
    yield* writeBinary(metadata, pages(), options.signal);
    assert(this.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during export');
  }
  async importBinary(
    input: Blob | Uint8Array | ReadableStream<Uint8Array>,
    options: { signal?: AbortSignal } = {},
  ) {
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    assert(
      !this.importRunning && !this.transactionRunning && this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    options.signal?.throwIfAborted();
    this.importRunning = true;
    const cancel = () => this.cancel();
    options.signal?.addEventListener('abort', cancel, { once: true });
    let acknowledged = false;
    try {
      const view = await this.rpc(
        'importBinary',
        [input],
        input instanceof ReadableStream ? [input] : [],
      );
      acknowledged = true;
      this.adopt(view);
      this.notifyRevision();
    } catch (error) {
      this.writeFailure = error;
      throw error;
    } finally {
      this.importRunning = false;
      options.signal?.removeEventListener('abort', cancel);
      if (acknowledged) this.pendingWrites--;
      this.emitPersistence();
    }
  }
  cancel() {
    this.transactionCancelled = true;
    this.worker.postMessage({ id: 0, method: 'cancel', args: [] });
  }
  async deleteStorage() {
    assert(!this.disposed, 'DISPOSED', 'Workbook is disposed');
    assert(!this.expired, 'DISPOSED', 'Transaction scope is disposed');
    assert(!this.batch && !this.transactionRunning, 'BUSY', 'Cannot delete during a transaction');
    assert(
      !this.importRunning && !this.generateRunning && this.pendingWrites === 0,
      'BUSY',
      'Another write operation is running',
    );
    assert(!this.deleting, 'BUSY', 'Workbook deletion is in progress');
    this.deleting = true;
    try {
      await this.rpc('clear');
      try {
        this.channel?.postMessage({ deleted: true });
      } catch {
        // Deletion is already committed even if this notification cannot be delivered.
      }
      this.dispose();
    } finally {
      this.deleting = false;
    }
  }
  calculation = {
    calculate: async () => {
      for (const sh of this.getSheets())
        for await (const _data of this.streamRange(sh.id, this.usedRange(sh.id))) {
        }
      return { revision: this.revision };
    },
    getState: () => ({ revision: this.revision, state: 'ready' as const }),
  };
  async close() {
    assert(!this.expired, 'DISPOSED', 'Transaction scope is disposed');
    assert(!this.batch && !this.transactionRunning, 'BUSY', 'Cannot close during a transaction');
    if (!this.disposed) {
      try {
        if (!this.failure) await this.rpc('close');
      } finally {
        this.dispose();
      }
    }
  }
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    for (const fn of this.disposeListeners) {
      try {
        notifyObserver(fn);
      } catch (error) {
        console.error(error);
      }
    }
    this.disposeListeners.clear();
    if (typeof window !== 'undefined') window.removeEventListener('focus', this.refreshOnVisible);
    if (typeof document !== 'undefined')
      document.removeEventListener('visibilitychange', this.refreshOnVisible);
    if (this.remoteRefreshTimer) clearTimeout(this.remoteRefreshTimer);
    this.channel?.close();
    this.worker.terminate();
    this.fail(new OpenSheetError('DISPOSED', 'Workbook closed'));
    this.cache.clear();
    this.seen.clear();
    this.listeners.clear();
    this.progressListeners.clear();
    this.dataListeners.clear();
    this.persistenceListeners.clear();
  }
}
export async function createWorkbook(
  options: Parameters<typeof createSnapshot>[0] = {},
  storage: WorkbookOptions = {},
) {
  return new Workbook(encodeWorkbookFile(createSnapshot(options)), storage).ready();
}
export async function openWorkbook(id: string, options: WorkbookOptions = {}) {
  return new Workbook(undefined, options, id).ready();
}
