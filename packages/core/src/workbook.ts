import { freeze } from 'immer';
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
import { BLOCK_ROWS, BLOCK_COLUMNS, ByteCache } from './storage.js';
import { blockAddress, type EngineView, type Progress } from './engine.js';
import {
  assert,
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
  type WorkbookSnapshot,
} from './types.js';
export interface WorkbookOptions {
  database?: string;
  budgetBytes?: number;
  cacheBytes?: number;
  viewportCacheBytes?: number;
  workerFactory?: () => Worker;
  operationTimeoutMs?: number;
  onProgress?: (progress: Progress) => void;
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
    }
  >();
  private nextId = 0;
  private disposed = false;
  private cache: ByteCache<RangeData>;
  private loading = new Map<string, Promise<void>>();
  private listeners = new Set<(commit: Commit) => void>();
  private dataListeners = new Set<() => void>();
  private batch?: Command[];
  private transactionRunning = false;
  private importRunning = false;
  private poisoned = false;
  private expired = false;
  private transactionCancelled = false;
  private readOnly = false;
  private writer = false;
  private channel?: BroadcastChannel;
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
      options.viewportCacheBytes === undefined ||
        (Number.isSafeInteger(options.viewportCacheBytes) && options.viewportCacheBytes >= 0),
      'INVALID_ARGUMENT',
      'Invalid viewport cache budget',
    );
    assert(
      options.workerFactory || typeof Worker !== 'undefined',
      'WORKER_UNAVAILABLE',
      'A browser Worker or injected workerFactory is required',
    );
    this.id = id;
    const cacheBudget = options.viewportCacheBytes ?? 8 * 1024 * 1024;
    const retryBytes = Math.min(512 * 1024, Math.floor(cacheBudget / 16));
    this.seen = new ByteCache(retryBytes);
    this.cache = new ByteCache(cacheBudget - retryBytes);
    this.worker =
      options.workerFactory?.() ??
      new Worker(new URL('./engine.worker.js', import.meta.url), { type: 'module' });
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
        options.onProgress?.(message.progress);
        return;
      }
      const task = this.pending.get(message.id);
      if (!task) return;
      this.pending.delete(message.id);
      clearTimeout(task.timer);
      if (message.error) task.reject(new OpenSheetError(message.error.code, message.error.message));
      else task.resolve(message.result);
    };
    this.worker.onerror = (event) => {
      this.fail(new OpenSheetError('WORKER_FAILED', event.message || 'Storage worker failed'));
    };
    this.worker.onmessageerror = () =>
      this.fail(new OpenSheetError('WORKER_FAILED', 'Invalid worker response'));
    this.initial = this.rpc('open', [
      id,
      snapshot,
      {
        database: options.database,
        budgetBytes: options.budgetBytes,
        cacheBytes: options.cacheBytes,
      },
    ])
      .then(({ view, writer }) => {
        if (this.disposed)
          throw new OpenSheetError('DISPOSED', 'Workbook closed during initialization');
        this.adopt(view);
        this.writer = writer;
        this.channel = new BroadcastChannel(`opensheet:${options.database ?? 'opensheet'}:${id}`);
        this.channel.onmessage = () => {
          if (!this.writer)
            void this.rpc('reload')
              .then((view) => {
                const previousRevision = this.revision;
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
              .catch((error) => this.fail(error));
        };
        return this;
      })
      .catch((error) => {
        this.dispose();
        throw error;
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
  }
  private rpc(method: string, args: unknown[] = [], transfer: Transferable[] = []): Promise<any> {
    assert(!this.disposed, 'DISPOSED', 'Workbook is disposed');
    if (this.failure) return Promise.reject(this.failure);
    const id = ++this.nextId;
    return new Promise((resolve, reject) => {
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
      this.pending.set(id, { resolve, reject, timer });
      try {
        this.worker.postMessage({ id, method, args }, transfer);
      } catch (error) {
        this.pending.delete(id);
        clearTimeout(timer);
        reject(error);
      }
    });
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
    this.readOnly = value;
  }
  getSheets() {
    return this.metadata.sheetOrder.map((id) => new Sheet(this, id));
  }
  getSheetById(id: string) {
    return this.metadata.sheets.some((s) => s.id === id) ? new Sheet(this, id) : undefined;
  }
  getSheetByName(name: string) {
    return this.getSheets().find((s) => s.name.toLowerCase() === name.toLowerCase());
  }
  sheetData(id: string) {
    const sh = this.metadata.sheets.find((s) => s.id === id);
    assert(sh, 'INVALID_ARGUMENT', 'Unknown sheet');
    return sh;
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
  async getCell(sheet: string, row: number, column: number) {
    return (
      await this.readRange(
        sheet,
        {
          startRow: row,
          endRow: row + 1,
          startColumn: column,
          endColumn: column + 1,
        },
        false,
      )
    ).cells[key(this.sheetData(sheet).rowOrder[row], this.sheetData(sheet).columnOrder[column])];
  }
  async value(sheet: string, row: number, column: number) {
    const data = await this.readRange(sheet, {
      startRow: row,
      endRow: row + 1,
      startColumn: column,
      endColumn: column + 1,
    });
    return (
      data.calculated[
        key(this.sheetData(sheet).rowOrder[row], this.sheetData(sheet).columnOrder[column])
      ] ?? null
    );
  }
  async display(sheet: string, row: number, column: number) {
    const data = await this.readRange(sheet, {
      startRow: row,
      endRow: row + 1,
      startColumn: column,
      endColumn: column + 1,
    });
    return (
      data.display[
        key(this.sheetData(sheet).rowOrder[row], this.sheetData(sheet).columnOrder[column])
      ] ?? ''
    );
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
              if (!this.disposed && revision === this.revision && data.revision === revision) {
                this.cache.set(a, data);
                for (const fn of this.dataListeners) fn();
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
    checkBounds(this.sheetData(sheet), range);
    const revision = this.revision;
    const width = range.endColumn - range.startColumn,
      rows = Math.max(1, Math.floor(16384 / width));
    for (let row = range.startRow; row < range.endRow; row += rows) {
      options.signal?.throwIfAborted();
      assert(
        revision === this.revision,
        'REVISION_CONFLICT',
        'Workbook changed during streaming read',
      );
      const data = await this.readRange(
        sheet,
        { ...range, startRow: row, endRow: Math.min(row + rows, range.endRow) },
        options.values ?? true,
      );
      assert(
        data.revision === revision,
        'REVISION_CONFLICT',
        'Workbook changed during streaming read',
      );
      yield data;
    }
  }
  private adopt(view: EngineView) {
    const wire = view as EngineView & { file?: WorkbookFile; styleIds?: string[] };
    const snapshot = wire.file ? decodeFileMetadata(wire.file) : this.view?.snapshot;
    assert(snapshot, 'CORRUPT_STORAGE', 'Missing workbook metadata');
    if (wire.file && wire.styleIds)
      snapshot.styles = Object.fromEntries(
        wire.styleIds.map((id, i) => [id, wire.file!.styles[i]]),
      );
    this.view = { ...view, snapshot: freeze({ ...snapshot, revision: view.revision }, true) };
    this.cache.clear();
    this.loading.clear();
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
  private emit(commit: Commit) {
    for (const fn of this.listeners) {
      try {
        fn(structuredClone(commit));
      } catch (error) {
        console.error(error);
      }
    }
  }
  private accept(result: { view: EngineView; commit?: Commit; duplicate?: boolean }) {
    if (result.duplicate) return result.commit;
    this.adopt(result.view);
    if (result.commit) {
      this.emit(result.commit);
      this.channel?.postMessage({ revision: this.revision });
    }
    return result.commit;
  }
  async execute(input: Command | CommandEnvelope): Promise<Commit | undefined> {
    assert(!this.importRunning, 'BUSY', 'Import is running');
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    if (this.batch) {
      try {
        assert(!('protocolVersion' in input), 'INVALID_ARGUMENT', 'Envelopes cannot be batched');
        this.batch.push(validateCommand(input));
        return;
      } catch (error) {
        this.poisoned = true;
        throw error;
      }
    }
    assert(!this.transactionRunning, 'BUSY', 'Use the scoped workbook inside a transaction');
    let signature = '';
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
        return seen.commit;
      }
      assert(input.baseRevision === this.revision, 'REVISION_CONFLICT', 'Workbook changed');
    }
    const commit = this.accept(await this.rpc('execute', [[input], 'edit']));
    if ('protocolVersion' in input) {
      this.seen.set(input.commandId, { signature, commit });
    }
    return commit;
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
    this.transactionRunning = true;
    this.transactionCancelled = false;
    const scoped = Object.create(this) as Workbook;
    scoped.batch = [];
    scoped.poisoned = false;
    scoped.transactionRunning = false;
    scoped.rpc = this.rpc.bind(this);
    scoped.dispose = () => {
      throw new OpenSheetError('INVALID_ARGUMENT', 'Cannot dispose a transaction scope');
    };
    try {
      const value = await fn(scoped);
      assert(!scoped.poisoned, 'INVALID_ARGUMENT', 'Transaction contains a failed command');
      assert(!this.transactionCancelled, 'ABORTED', 'Transaction cancelled');
      if (scoped.batch.length)
        this.accept(await this.rpc('execute', [scoped.batch, options.label ?? 'transaction']));
      return value;
    } finally {
      scoped.batch = undefined;
      scoped.expired = true;
      this.transactionRunning = false;
    }
  }
  async undo() {
    assert(!this.batch && !this.transactionRunning, 'BUSY', 'Cannot undo during a transaction');
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    return !!this.accept(await this.rpc('history', [false]));
  }
  async redo() {
    assert(!this.batch && !this.transactionRunning, 'BUSY', 'Cannot redo during a transaction');
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    return !!this.accept(await this.rpc('history', [true]));
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
  search(sheet: string, query: string, after = -1) {
    return this.rpc('search', [sheet, query, after]) as Promise<{
      match?: { row: number; column: number };
      count: number;
    }>;
  }
  async generate(count: number, budget = this.view.budget) {
    assert(!this.isReadOnly, 'READ_ONLY', 'Workbook is read only');
    this.adopt(await this.rpc('generate', [count, budget]));
    this.channel?.postMessage({ revision: this.revision });
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
      !this.importRunning && !this.transactionRunning,
      'BUSY',
      'Another write operation is running',
    );
    options.signal?.throwIfAborted();
    this.importRunning = true;
    const cancel = () => this.cancel();
    options.signal?.addEventListener('abort', cancel, { once: true });
    try {
      this.adopt(
        await this.rpc('importJSON', [input], input instanceof ReadableStream ? [input] : []),
      );
      this.channel?.postMessage({ revision: this.revision });
    } finally {
      this.importRunning = false;
      options.signal?.removeEventListener('abort', cancel);
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
      { file, ids } = await this.exportMetadata(revision);
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
      !this.importRunning && !this.transactionRunning,
      'BUSY',
      'Another write operation is running',
    );
    options.signal?.throwIfAborted();
    this.importRunning = true;
    const cancel = () => this.cancel();
    options.signal?.addEventListener('abort', cancel, { once: true });
    try {
      this.adopt(
        await this.rpc('importBinary', [input], input instanceof ReadableStream ? [input] : []),
      );
      this.channel?.postMessage({ revision: this.revision });
    } finally {
      this.importRunning = false;
      options.signal?.removeEventListener('abort', cancel);
    }
  }
  cancel() {
    this.transactionCancelled = true;
    this.worker.postMessage({ id: 0, method: 'cancel', args: [] });
  }
  async deleteStorage() {
    await this.rpc('clear');
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
    this.channel?.close();
    this.worker.terminate();
    this.fail(new OpenSheetError('DISPOSED', 'Workbook closed'));
    this.cache.clear();
    this.seen.clear();
    this.listeners.clear();
    this.dataListeners.clear();
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
