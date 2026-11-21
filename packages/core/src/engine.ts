import { matchesFilter } from './filter.js';
import { yieldToTask } from './task-yield.js';
import { validateSearchArguments } from './search-validation.js';
import { validateNavigationArguments } from './navigation-validation.js';
import { numericSeries } from './numeric-series.js';
import { calendarSeries } from './calendar-series.js';
import { createTileOffsetter, tileCell, validatePasteTile, type PasteTile } from './paste-tile.js';
import { compareSortKeys, sortColumns, sortKey } from './sort-values.js';
import { AxisIndex } from './axis-index.js';
import { overlay } from './copy-on-write.js';
import { contains, intersects } from './address.js';
import { parseFormula, type AST } from '@opensheetjs/formula';
import { decodeBlock, encodeBlock, ColumnPage, type EncodedBlock } from './block-codec.js';
import {
  decodeFileMetadata,
  encodeWorkbookFile,
  decodeFileCells,
  workbookFileSchema,
  type WorkbookFile,
} from './file-codec.js';
import { crc32, gzipPage, unpackPage, readBinary } from './binary-file.js';
import { decodeVerifiedBinaryPage } from './binary-verify.js';
import { parseWorkbookJSON, type JSONSource } from './json-stream.js';
import { downgradeValidationMetadata } from './validation-file-migration.js';
import { checkStagedValidation } from './validation-batch.js';
import { validationCandidates } from './validation-candidates.js';
import type { ValidationEntry } from './validation-model.js';
import {
  conditionalViewportWithFormulas,
  validateConditionalRules,
  type ConditionalRule,
} from './conditional-format.js';
import { abortable } from './verification-yield.js';
import { AsyncCalculation } from './async-calculation.js';
import { selectiveCopyCell } from './selective-copy.js';
import { reduceCommand, validateCommand } from './commands.js';
import {
  applyPastedCell,
  createCellFormulaTransformer,
  structuralRangeReferences,
  offsetCellFormula,
  structuralReferences,
  styleId,
  type PasteCell,
} from './commands/helpers.js';
import { formatValue } from './format.js';
import { checkBounds, validateJSONValue, validateSnapshot } from './model.js';
import {
  BLOCK_ROWS,
  BLOCK_COLUMNS,
  ByteCache,
  Database,
  encodedBytes,
  recordBytes,
} from './storage.js';
import {
  assert,
  LIMITS,
  OpenSheetError,
  key,
  uid,
  scalarValue,
  type CellRecord,
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
function groupAxis(order: number[], start: number, end: number, size: number) {
  const result = new Map<number, { index: number; id: number }[]>();
  for (let index = start; index < end; index++) {
    const id = order[index],
      group = Math.floor(id / size);
    let entries = result.get(group);
    if (!entries) result.set(group, (entries = []));
    entries.push({ index, id });
  }
  return result;
}
function changesStyles(command: Command): boolean {
  return (
    command.type === 'core.cells.style' ||
    ((command.type === 'core.cells.paste' || command.type === 'core.cells.pasteTile') &&
      (command.payload as { cells: { style?: unknown }[] }).cells.some(
        (cell) => cell.style !== undefined,
      ))
  );
}
export type Block = { cells: Record<string, CellRecord>; dependencies: Record<string, unknown> };
type StagedReadPage = { address?: string; page?: ColumnPage };
const CONDITIONAL_READ_PAGE_BUDGET = 1024 * 1024;
const SEARCH_YIELD_TEXT_UNITS = 1024 * 1024;
export type Pointer = {
  id: string;
  bytes: number;
  count: number;
  endRow: number;
  endColumn: number;
  formulas?: string[];
  styles?: string[];
  rowMask?: string;
  columnMask?: number;
};
type Head = {
  engineFormatVersion?: number;
  directory?: Record<string, string>;
  meta: string;
  blocks: Record<string, Pointer>;
  revision: number;
  past: string[];
  future: string[];
  historyBytes: number;
  budget: number;
};
type Metadata = {
  importRevision?: number;
  snapshot: WorkbookSnapshot;
  filtered: Record<string, number[]>;
};
type StoredMetadata = Omit<Metadata, 'snapshot'> & { snapshot: WorkbookFile; styleIds: string[] };
const encodeMetadata = (metadata: Metadata): StoredMetadata => ({
  ...metadata,
  snapshot: encodeWorkbookFile(metadata.snapshot),
  styleIds: Object.keys(metadata.snapshot.styles),
});
type History = {
  before: string;
  after: string;
  changes: Array<[string, Pointer | null, Pointer | null]>;
  bytes: number;
  ranges: Selection[];
};
export type EngineView = {
  metadataId: string;
  revision: number;
  importRevision?: number;
  snapshot: WorkbookSnapshot;
  filtered: Record<string, number[]>;
  used: Record<string, Rect>;
  canUndo: boolean;
  canRedo: boolean;
  bytes: number;
  cacheBytes: number;
  budget: number;
};
export type Progress = { phase: string; completed: number; total?: number; elapsedMs?: number };
export type ReadTiming = {
  cacheHits: number;
  cacheMisses: number;
  absentPages: number;
  storageMs: number;
  decodeMs: number;
  inflateMs: number;
  pageDecodeMs: number;
};
export const blockAddress = (sheet: string, row: number, column: number) =>
  `${sheet}/${Math.floor(row / BLOCK_ROWS)}/${Math.floor(column / BLOCK_COLUMNS)}`;
const prefixRange = (id: string) => IDBKeyRange.bound(`${id}/`, `${id}/￿`);
export class Engine {
  private head!: Head;
  private metadata!: Metadata;
  private usedBytes = 0;
  private allocations: string[] = [];
  private allocated = new Set<string>();
  private stagingBytes = 0;
  private pins = new Set<string>();
  private axes = new WeakMap<number[], AxisIndex>();
  private axis(ids: number[]) {
    let index = this.axes.get(ids);
    if (!index) {
      index = new AxisIndex(ids);
      this.axes.set(ids, index);
    }
    return index;
  }
  private indices = new WeakMap<object, { rows: AxisIndex; cols: AxisIndex }>();
  private exportAddresses: string[] | undefined;
  private searchResults = new WeakMap<
    Head,
    { sheetId: string; query: string; positions: Uint32Array }
  >();
  private compressionQueue = new Set<string>();
  private directoryEntries = new Map<string, Record<string, Pointer>>();
  private observedHeadSignature?: string;
  private changed = new Set<string>();
  private references = new Map<string, number>();
  private histories = new Map<string, History>();
  private used: Record<string, Rect> = {};
  private sizes = new Map<string, number>();
  private cache: ByteCache<ColumnPage>;
  private cancelled = false;
  private importRead?: AbortController;
  private poisoned = false;
  private writer = false;
  private release?: () => void;
  private calculator: AsyncCalculation;
  private seen: ByteCache<{ signature: string; commit?: Commit }>;
  constructor(
    private db: Database,
    readonly id: string,
    cacheBytes = 64 * 1024 * 1024,
    private progress: (p: Progress) => void = () => {},
  ) {
    const retryBytes = Math.min(2 * 1024 * 1024, Math.floor(cacheBytes / 16));
    this.seen = new ByteCache(retryBytes);
    const calculationBytes = Math.min(4 * 1024 * 1024, Math.floor(cacheBytes / 8));
    this.cache = new ByteCache(Math.max(0, cacheBytes - calculationBytes - retryBytes));
    this.calculator = new AsyncCalculation(
      (s, r, c) => this.cell(s, r, c),
      (name) =>
        this.metadata.snapshot.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())?.id,
      calculationBytes,
      async (sheet, row, column) => {
        const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheet)!;
        return (
          await this.block(
            blockAddress(
              sheet,
              this.metadata.snapshot.sheets.find((s) => s.id === sheet)!.rowOrder[row],
              this.metadata.snapshot.sheets.find((s) => s.id === sheet)!.columnOrder[column],
            ),
          )
        ).dependencies[key(sh.rowOrder[row], sh.columnOrder[column])] as AST | null | undefined;
      },
      () => this.metadata.snapshot.dateSystem,
    );
  }
  get isWriter() {
    return this.writer;
  }
  async lock() {
    try {
      const locks = typeof navigator !== 'undefined' ? navigator.locks : undefined;
      const request = locks?.request;
      if (typeof request === 'function') {
        await new Promise<void>((resolve, reject) => {
          void request
            .call(
              locks,
              `opensheet:${this.db.db.name}:${this.id}`,
              { ifAvailable: true },
              async (lock) => {
                this.writer = !!lock;
                resolve();
                if (lock)
                  await new Promise<void>((release) => {
                    this.release = release;
                  });
              },
            )
            .catch(reject);
        });
      } else if (typeof window === 'undefined' && typeof self === 'undefined') this.writer = true;
      else
        throw new OpenSheetError(
          'WEB_LOCKS_UNAVAILABLE',
          'Web Locks is required for safe multi-tab editing; open this app over HTTPS or localhost in a supported browser',
        );
    } catch (error) {
      const name = (error as { name?: unknown })?.name;
      if (name === 'SecurityError' || name === 'NotAllowedError')
        throw new OpenSheetError(
          'WEB_LOCKS_UNAVAILABLE',
          'Web Locks access was denied; open this app over HTTPS or localhost in a supported browser',
        );
      throw error;
    }
  }
  async tryAcquireWriteLock() {
    if (this.writer) return { writer: true, view: this.view() };
    await this.lock();
    if (!this.writer) return { writer: false };
    try {
      await this.reload();
      await this.recover();
      return { writer: true, view: this.view() };
    } catch (error) {
      this.release?.();
      this.release = undefined;
      this.writer = false;
      throw error;
    }
  }
  private check() {
    assert(!this.cancelled, 'ABORTED', 'Operation cancelled');
  }
  begin() {
    assert(!this.poisoned, 'WORKER_FAILED', 'Reopen the workbook after an uncertain commit');
    this.cancelled = false;
  }
  cancel() {
    this.cancelled = true;
    this.importRead?.abort(new OpenSheetError('ABORTED', 'Operation cancelled'));
  }
  async open(
    input?: WorkbookSnapshot | WorkbookFile,
    budget = 256 * 1024 * 1024,
    preserveOrphans = true,
  ) {
    assert(
      Number.isSafeInteger(budget) && budget > 0,
      'INVALID_ARGUMENT',
      'Storage budget must be a positive integer',
    );
    const snapshot =
      input &&
      ('sheetOrder' in input
        ? input
        : decodeFileMetadata({
            ...input,
            sheets: input.sheets.map((s) => ({ ...s, blocks: [] })),
          }));
    await this.lock();
    if (!this.writer) {
      await this.loadReadOnly(true);
      return this.view();
    }
    const storedHead = await this.db.get<Head>('heads', this.id);
    const existing = storedHead !== undefined ? await this.restoreHead(storedHead) : undefined;
    if (existing) {
      this.head = existing;
      this.metadata = await this.readMetadata(this.head.meta);
      assert(this.metadata, 'CORRUPT_STORAGE', 'Workbook metadata is missing');
      this.checkDirectorySheets();
      // Visibility is derived from cells; scalar history entries can share metadata.
      await this.filters(this.head, this.metadata);
      await this.recover();
    } else {
      assert(snapshot && this.writer, 'INVALID_ARGUMENT', 'Workbook does not exist or is locked');
      // Existing records without a head may be an interrupted creation or lost user data.
      if (preserveOrphans) {
        const range = prefixRange(this.id);
        const [records, sizes, scratch] = await Promise.all([
          this.db.hasAny('records', range),
          this.db.hasAny('sizes', range),
          this.db.hasAny('scratch', range),
        ]);
        assert(
          !records && !sizes && !scratch,
          'CORRUPT_STORAGE',
          'Workbook head is missing but stored data remains; restore a backup',
        );
      }
      await this.db.write('records', [], [prefixRange(this.id)]);
      await this.db.write('scratch', [], [prefixRange(this.id)]);
      this.head = {
        meta: '',
        blocks: {},
        revision: snapshot.revision,
        past: [],
        future: [],
        historyBytes: 0,
        budget,
      };
      this.usedBytes = recordBytes(this.id, this.storedHead(this.head));
      this.metadata = {
        snapshot: { ...snapshot, sheets: snapshot.sheets.map((s) => ({ ...s, cells: {} })) },
        filtered: {},
      };
      if (input && !('sheetOrder' in input)) await this.importJSON(input, budget);
      else await this.replace(snapshot, budget);
    }
    return this.view();
  }
  private async liveIds(head = this.head) {
    const ids = new Set([
      head.meta,
      ...Object.values(head.blocks).map((p) => p.id),
      ...Object.values(head.directory ?? {}),
      ...head.past,
      ...head.future,
    ]);
    for (const id of [...head.past, ...head.future]) {
      const h = await this.readHistory(id);
      ids.add(h.before);
      ids.add(h.after);
      for (const [, a, b] of h.changes) {
        if (a) ids.add(a.id);
        if (b) ids.add(b.id);
      }
    }
    return ids;
  }
  private async liveBytes() {
    let total = recordBytes(this.id, this.storedHead(this.head));
    const ids = [...(await this.liveIds())];
    const sizes = await this.db.recordSizes(ids);
    ids.forEach((id, i) => {
      total += sizes[i] + recordBytes(id, sizes[i]);
    });
    return total;
  }
  private retain(id: string, delta: number) {
    if (!id) return;
    const count = (this.references.get(id) ?? 0) + delta;
    assert(count >= 0, 'CORRUPT_STORAGE', 'Negative record reference count');
    if (count) this.references.set(id, count);
    else this.references.delete(id);
  }
  private retainHistory(id: string, h: History, delta: number) {
    this.retain(id, delta);
    this.retain(h.before, delta);
    this.retain(h.after, delta);
    for (const [, a, b] of h.changes) {
      if (a) this.retain(a.id, delta);
      if (b) this.retain(b.id, delta);
    }
  }
  private validPointer(value: unknown): value is Pointer {
    if (!value || typeof value !== 'object' || Array.isArray(value)) return false;
    const pointer = value as Pointer;
    return (
      Object.keys(pointer).every((key) =>
        [
          'id',
          'bytes',
          'count',
          'endRow',
          'endColumn',
          'formulas',
          'styles',
          'rowMask',
          'columnMask',
        ].includes(key),
      ) &&
      typeof pointer.id === 'string' &&
      pointer.id.startsWith(`${this.id}/`) &&
      Number.isSafeInteger(pointer.bytes) &&
      pointer.bytes > 0 &&
      Number.isInteger(pointer.count) &&
      pointer.count >= 0 &&
      pointer.count <= BLOCK_ROWS * BLOCK_COLUMNS &&
      Number.isInteger(pointer.endRow) &&
      pointer.endRow > 0 &&
      pointer.endRow <= LIMITS.rows &&
      Number.isInteger(pointer.endColumn) &&
      pointer.endColumn > 0 &&
      pointer.endColumn <= LIMITS.columns &&
      (pointer.rowMask === undefined ||
        (typeof pointer.rowMask === 'string' && /^[\da-f]{1,16}$/i.test(pointer.rowMask))) &&
      (pointer.columnMask === undefined ||
        (Number.isInteger(pointer.columnMask) &&
          pointer.columnMask >= 0 &&
          pointer.columnMask <= 0xffffffff)) &&
      (pointer.formulas === undefined ||
        (Array.isArray(pointer.formulas) &&
          pointer.formulas.every((name) => typeof name === 'string'))) &&
      (pointer.styles === undefined ||
        (Array.isArray(pointer.styles) &&
          pointer.styles.every((style) => typeof style === 'string')))
    );
  }
  private async readHistory(id: string): Promise<History> {
    const history = await this.db.get<History>('records', id);
    const owned = (value: unknown) => typeof value === 'string' && value.startsWith(`${this.id}/`);
    const pointer = (value: unknown) => value === null || this.validPointer(value);
    const address = (value: unknown) => {
      if (typeof value !== 'string') return false;
      const match = /^[\w-]+\/(0|[1-9]\d*)\/(0|[1-9]\d*)$/.exec(value);
      return (
        !!match &&
        Number(match[1]) <= Math.floor(0xffffffff / BLOCK_ROWS) &&
        Number(match[2]) <= Math.floor(0xffffffff / BLOCK_COLUMNS)
      );
    };
    assert(
      history &&
        typeof history === 'object' &&
        !Array.isArray(history) &&
        Object.keys(history).every((key) =>
          ['before', 'after', 'changes', 'bytes', 'ranges'].includes(key),
        ) &&
        owned(history.before) &&
        owned(history.after) &&
        Number.isSafeInteger(history.bytes) &&
        history.bytes >= 0 &&
        history.bytes <= 32 * 1024 * 1024 &&
        Array.isArray(history.changes) &&
        history.changes.every(
          (change) =>
            Array.isArray(change) &&
            change.length === 3 &&
            address(change[0]) &&
            (change[1] !== null || change[2] !== null) &&
            pointer(change[1]) &&
            pointer(change[2]),
        ) &&
        new Set(history.changes.map((change) => change[0])).size === history.changes.length &&
        Array.isArray(history.ranges) &&
        history.ranges.every(
          (range) =>
            range &&
            typeof range === 'object' &&
            !Array.isArray(range) &&
            Object.keys(range).every((key) =>
              ['sheetId', 'startRow', 'endRow', 'startColumn', 'endColumn'].includes(key),
            ) &&
            typeof range.sheetId === 'string' &&
            Number.isInteger(range.startRow) &&
            range.startRow >= 0 &&
            Number.isInteger(range.endRow) &&
            range.endRow > range.startRow &&
            range.endRow <= LIMITS.rows &&
            Number.isInteger(range.startColumn) &&
            range.startColumn >= 0 &&
            Number.isInteger(range.endColumn) &&
            range.endColumn > range.startColumn &&
            range.endColumn <= LIMITS.columns,
        ),
      'CORRUPT_STORAGE',
      'Invalid workbook history; open with a compatible version or restore a backup',
    );
    return history;
  }
  private async initializeReferences() {
    this.references.clear();
    this.histories.clear();
    this.retain(this.head.meta, 1);
    for (const id of Object.values(this.head.directory ?? {})) this.retain(id, 1);
    for (const p of Object.values(this.head.blocks)) this.retain(p.id, 1);
    let pastBytes = 0;
    const checkedMetadata = new Set([this.head.meta]);
    for (const [index, id] of [...this.head.past, ...this.head.future].entries()) {
      const h = await this.readHistory(id);
      for (const metadataId of [h.before, h.after]) {
        if (checkedMetadata.has(metadataId)) continue;
        // Refuse to open a history that this engine could neither restore nor safely prune.
        const stored = await this.readMetadataEnvelope(metadataId);
        const { snapshot } = stored;
        // Persisted metadata has no cell pages. Validate its compressed shape
        // without allocating the expanded row/column orders of each revision.
        assert(
          Array.isArray(snapshot.sheets) &&
            snapshot.sheets.every(
              (sheet) => Array.isArray(sheet?.blocks) && sheet.blocks.length === 0,
            ),
          'CORRUPT_STORAGE',
          'Invalid historical workbook metadata',
        );
        validateJSONValue(snapshot);
        assert(
          workbookFileSchema.safeParse(snapshot).success,
          'CORRUPT_STORAGE',
          'Invalid historical workbook metadata; open with a compatible version or restore a backup',
        );
        this.validateStoredStyleIds(stored);
        checkedMetadata.add(metadataId);
      }
      if (index < this.head.past.length) pastBytes += h.bytes;
      this.histories.set(id, h);
      this.retainHistory(id, h, 1);
    }
    assert(
      pastBytes === this.head.historyBytes,
      'CORRUPT_STORAGE',
      'Workbook history accounting differs; restore a backup',
    );
    await this.checkHistoricalDirectories();
  }

  private async checkHistoricalDirectories() {
    // Reconstruct both branches from the current directory. Counts avoid scanning
    // every page for every revision while still checking unchanged pages when
    // a history entry changes the metadata.
    for (const redo of [false, true]) {
      const historyIds = redo ? this.head.future : this.head.past;
      if (!historyIds.length) continue;
      const blocks = { ...this.head.blocks };
      const sheets = new Map<string, number>();
      const styles = new Map<string, number>();
      const adjust = (counts: Map<string, number>, key: string, delta: number) => {
        const count = (counts.get(key) ?? 0) + delta;
        if (count) counts.set(key, count);
        else counts.delete(key);
      };
      const account = (address: string, pointer: Pointer, delta: number) => {
        adjust(sheets, address.slice(0, address.indexOf('/')), delta);
        for (const style of pointer.styles ?? []) adjust(styles, style, delta);
      };
      for (const [address, pointer] of Object.entries(blocks)) account(address, pointer, 1);
      let metadataId = this.head.meta;
      let validSheets = new Set(this.metadata.snapshot.sheets.map((sheet) => sheet.id));
      let validStyles = new Set(Object.keys(this.metadata.snapshot.styles));
      for (let i = historyIds.length - 1; i >= 0; i--) {
        const history = this.histories.get(historyIds[i])!;
        for (const [address, before, after] of history.changes) {
          if (Object.hasOwn(blocks, address)) account(address, blocks[address], -1);
          const pointer = redo ? after : before;
          if (pointer) {
            blocks[address] = pointer;
            account(address, pointer, 1);
          } else delete blocks[address];
        }
        const targetId = redo ? history.after : history.before;
        if (targetId !== metadataId) {
          const stored = await this.readMetadataEnvelope(targetId);
          validSheets = new Set(stored.snapshot.sheets.map((sheet) => sheet.id));
          validStyles = new Set(stored.styleIds);
          metadataId = targetId;
        }
        for (const id of sheets.keys())
          assert(
            validSheets.has(id),
            'CORRUPT_STORAGE',
            'Block directory references an unknown sheet',
          );
        for (const id of styles.keys())
          assert(
            validStyles.has(id),
            'CORRUPT_STORAGE',
            'Block directory references an unknown style',
          );
      }
    }
  }
  private async recover() {
    await this.initializeReferences();
    for (const id of this.references.keys()) this.compressionQueue.add(id);
    const live = await this.liveIds();
    const liveList = [...live];
    const ledger = await this.db.recordSizes(liveList);
    this.usedBytes = recordBytes(this.id, this.storedHead(this.head));
    this.sizes.clear();
    liveList.forEach((id, i) => {
      const bytes = ledger[i] + recordBytes(id, ledger[i]);
      this.usedBytes += bytes;
      this.sizes.set(id, bytes);
    });
    // Validate the ledger before any reclamation; a late malformed entry must not
    // leave a partially cleaned database when recovery reports failure.
    for await (const batch of this.db.scan<number>('sizes', prefixRange(this.id))) {
      for (const [, value] of batch)
        assert(
          Number.isSafeInteger(value) && value > 0,
          'CORRUPT_STORAGE',
          'Invalid storage ledger; restore a backup',
        );
    }
    for await (const batch of this.db.scanKeys('sizes', prefixRange(this.id))) {
      const dead: string[] = [];
      for (const k of batch) {
        const id = String(k);
        if (!live.has(id)) dead.push(id);
      }
      if (dead.length) await this.db.write('records', [], dead);
    }
    await this.db.write('scratch', [], [prefixRange(this.id)]);
  }
  async reload() {
    if (!this.writer) {
      await this.loadReadOnly(false);
      return this.view();
    }
    const storedHead = await this.db.get<Head>('heads', this.id);
    assert(storedHead, 'CORRUPT_STORAGE', 'Workbook head is missing; restore a backup');
    this.head = await this.restoreHead(storedHead);
    this.metadata = await this.readMetadata(this.head.meta);
    this.checkDirectorySheets();
    this.exportAddresses = undefined;
    this.cache.clear();
    this.calculator.clear();
    this.usedBytes = await this.liveBytes();
    await this.initializeReferences();
    this.used = {};
    return this.view();
  }
  async reloadIfChanged() {
    const current = await this.db.get<Head>('heads', this.id);
    assert(current, 'CORRUPT_STORAGE', 'Workbook head is missing; restore a backup');
    if (JSON.stringify(current) === this.observedHeadSignature) return undefined;
    return this.reload();
  }
  async headExists() {
    return (await this.db.get<Head>('heads', this.id)) !== undefined;
  }
  private async loadReadOnly(deriveFilters: boolean) {
    const started = Date.now();
    for (let attempt = 0; attempt < 64; attempt++) {
      const stored = await this.db.get<Head>('heads', this.id);
      assert(stored, 'CORRUPT_STORAGE', 'Workbook head is missing; restore a backup');
      const signature = JSON.stringify(stored);
      let failed = false,
        failure: unknown;
      try {
        this.head = await this.restoreHead(stored);
        this.metadata = await this.readMetadata(this.head.meta);
        this.checkDirectorySheets();
        this.exportAddresses = undefined;
        this.cache.clear();
        this.calculator.clear();
        this.used = {};
        if (deriveFilters) await this.filters(this.head, this.metadata);
        this.usedBytes = await this.liveBytes();
        await this.initializeReferences();
      } catch (error) {
        failed = true;
        failure = error;
      }
      const current = await this.db.get<Head>('heads', this.id);
      if (JSON.stringify(current) !== signature) {
        if (attempt === 63 || Date.now() - started >= 5000)
          throw new OpenSheetError(
            'REVISION_CONFLICT',
            'Workbook changed while opening; try again',
          );
        await new Promise<void>((resolve) => setTimeout(resolve, 10));
        continue;
      }
      if (failed) throw failure;
      this.observedHeadSignature = signature;
      return;
    }
  }
  private async allocate(value: unknown, budget = this.head.budget) {
    this.check();
    const id = `${this.id}/${uid('record')}`;
    const rawBytes = recordBytes(id, value);
    const bytes = rawBytes + recordBytes(id, rawBytes);
    assert(
      this.usedBytes + this.stagingBytes + bytes <= budget,
      'STORAGE_BUDGET',
      'Application storage budget exceeded (including staging and history)',
    );
    await this.db.write('records', [[id, value]]);
    this.usedBytes += bytes;
    this.sizes.set(id, bytes);
    this.allocations.push(id);
    this.allocated.add(id);
    return id;
  }
  private async readMetadataEnvelope(id: string): Promise<StoredMetadata> {
    const stored = await this.db.get<StoredMetadata>('records', id);
    assert(stored, 'CORRUPT_STORAGE', 'Workbook metadata is missing');
    assert(
      typeof stored === 'object' &&
        !Array.isArray(stored) &&
        Object.keys(stored).every((key) =>
          ['snapshot', 'styleIds', 'filtered', 'importRevision'].includes(key),
        ),
      'CORRUPT_STORAGE',
      'Unknown stored workbook metadata; open with a compatible version or restore a backup',
    );
    const version = stored.snapshot?.schemaVersion;
    assert(
      Number.isSafeInteger(version) && version > 0,
      'CORRUPT_STORAGE',
      'Invalid stored workbook schema',
    );
    assert(
      version === 3,
      'UNSUPPORTED_VERSION',
      `Stored workbook schema ${version} is unsupported; use a compatible version`,
    );
    return stored;
  }
  private validateStoredStyleIds(stored: StoredMetadata): void {
    assert(
      Array.isArray(stored.snapshot.styles) &&
        Array.isArray(stored.styleIds) &&
        stored.styleIds.length === stored.snapshot.styles.length &&
        new Set(stored.styleIds).size === stored.styleIds.length &&
        stored.styleIds.every(
          (id) =>
            typeof id === 'string' &&
            id.length <= 100 &&
            /^[-\w]+$/.test(id) &&
            !['__proto__', 'prototype', 'constructor'].includes(id),
        ),
      'CORRUPT_STORAGE',
      'Invalid stored style identities',
    );
  }
  private async readMetadata(id: string): Promise<Metadata> {
    const stored = await this.readMetadataEnvelope(id);
    const snapshot = decodeFileMetadata(stored.snapshot);
    this.validateStoredStyleIds(stored);
    snapshot.styles = Object.fromEntries(
      stored.styleIds.map((id, i) => [id, stored.snapshot.styles[i]]),
    );
    const { styleIds, ...metadata } = stored;
    return { ...metadata, snapshot };
  }
  private compareAddress(a: string, b: string) {
    const [sa, ra, ca] = a.split('/'),
      [sb, rb, cb] = b.split('/');
    return sa < sb ? -1 : sa > sb ? 1 : +ra - +rb || +ca - +cb;
  }
  exportIndex(revision: number, after?: string, count = 32, sheetId?: string): string[] {
    assert(revision === this.head.revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    assert(
      Number.isInteger(count) && count > 0 && count <= 64,
      'INVALID_ARGUMENT',
      'Invalid page batch',
    );
    const addresses = (this.exportAddresses ??= Object.keys(this.head.blocks).sort((a, b) =>
      this.compareAddress(a, b),
    ));
    let low = 0,
      high = addresses.length;
    const before = (address: string) =>
      after
        ? this.compareAddress(address, after) <= 0
        : sheetId
          ? address.split('/')[0] < sheetId
          : false;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (before(addresses[mid])) low = mid + 1;
      else high = mid;
    }
    const output: string[] = [];
    for (let i = low; i < addresses.length && output.length < count; i++) {
      const address = addresses[i];
      if (sheetId && !address.startsWith(`${sheetId}/`)) break;
      output.push(address);
    }
    return output;
  }
  exportStyles(revision: number) {
    assert(revision === this.head.revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    const used = new Set(
      Object.values(this.head.blocks).flatMap((pointer) => pointer.styles ?? []),
    );
    return Object.keys(this.metadata.snapshot.styles).filter((id) => used.has(id));
  }
  async exportBlock(address: string, revision: number): Promise<EncodedBlock> {
    assert(revision === this.head.revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    const pointer = this.head.blocks[address];
    assert(pointer, 'INVALID_ARGUMENT', 'Unknown export block');
    const encoded = await this.db.get<EncodedBlock>('records', pointer.id);
    this.check();
    assert(encoded, 'CORRUPT_STORAGE', 'Export block is missing');
    assert(revision === this.head.revision, 'REVISION_CONFLICT', 'Workbook changed during export');
    return this.inflate(encoded);
  }
  private async inflate(bytes: Uint8Array) {
    if (bytes[0] !== 79 || bytes[1] !== 83 || bytes[2] !== 71) return bytes;
    assert(bytes.length >= 12 && bytes[3] === 3, 'CORRUPT_STORAGE', 'Invalid compressed page');
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const page = await unpackPage(bytes.subarray(12), 1, view.getUint32(4, true));
    assert(
      crc32(page) === view.getUint32(8, true),
      'CORRUPT_STORAGE',
      'Stored page checksum differs',
    );
    return page;
  }
  get hasPendingCompaction() {
    return this.compressionQueue.size > 0;
  }
  async compact(count = 8) {
    if (!this.writer) return;
    for (const id of [...this.compressionQueue].slice(0, count)) {
      this.compressionQueue.delete(id);
      if (!this.references.has(id)) continue;
      const bytes = await this.db.get<unknown>('records', id);
      if (!(bytes instanceof Uint8Array) || bytes.length < 512 || bytes[2] !== 80) continue;
      const compressed = await gzipPage(bytes);
      if (compressed.codec === 0 || compressed.bytes.length + 12 > bytes.length * 0.9) continue;
      const record = new Uint8Array(compressed.bytes.length + 12),
        view = new DataView(record.buffer);
      record.set([79, 83, 71, 3]);
      view.setUint32(4, bytes.length, true);
      view.setUint32(8, crc32(bytes), true);
      record.set(compressed.bytes, 12);
      const raw = recordBytes(id, record),
        size = raw + recordBytes(id, raw),
        old = this.sizes.get(id)!;
      // Include replacement staging before the atomic record/ledger transaction.
      if (this.usedBytes + size > this.head.budget) continue;
      await this.db.write('records', [[id, record]]);
      this.sizes.set(id, size);
      this.usedBytes += size - old;
    }
    return {
      bytes: this.usedBytes,
      cacheBytes: this.cache.bytes + this.calculator.cacheBytes + this.seen.bytes,
    };
  }
  private async collect(candidates: Iterable<string>) {
    if (this.poisoned) return;
    const dead = [...new Set(candidates)].filter((id) => !this.references.has(id));
    if (dead.length) await this.db.write('records', [], dead);
    for (const id of dead) {
      this.usedBytes -= this.sizes.get(id) ?? 0;
      this.sizes.delete(id);
      this.cache.delete(id);
    }
  }
  private async cleanup(candidates: Iterable<string>) {
    // Failed reclamation remains charged and is retried on reopen. It must not
    // replace either a successful commit or the original operation error.
    try {
      await this.collect(candidates);
    } catch {
      /* Keep unreclaimed bytes charged. */
    }
  }
  private bucket(address: string) {
    let hash = 2166136261;
    for (const c of address) hash = Math.imul(hash ^ c.charCodeAt(0), 16777619);
    return String((hash >>> 0) % 256);
  }
  private storedHead(head: Head) {
    const { blocks, ...rest } = head;
    return { ...rest, directory: head.directory ?? {} };
  }
  private checkDirectorySheets(head = this.head, metadata = this.metadata) {
    const sheetIds = new Set(metadata.snapshot.sheets.map((sheet) => sheet.id));
    for (const address in head.blocks) {
      if (!Object.hasOwn(head.blocks, address)) continue;
      const pointer = head.blocks[address];
      assert(
        sheetIds.has(address.slice(0, address.indexOf('/'))),
        'CORRUPT_STORAGE',
        'Block directory references an unknown sheet',
      );
      assert(
        pointer.styles === undefined ||
          pointer.styles.every((id) => Object.hasOwn(metadata.snapshot.styles, id)),
        'CORRUPT_STORAGE',
        'Block directory references an unknown style',
      );
    }
  }
  private async restoreHead(head: Head) {
    assert(
      head && typeof head === 'object' && !Array.isArray(head),
      'CORRUPT_STORAGE',
      'Invalid workbook head',
    );
    if (Object.hasOwn(head, 'engineFormatVersion')) {
      assert(
        Number.isSafeInteger(head.engineFormatVersion) && head.engineFormatVersion! > 0,
        'CORRUPT_STORAGE',
        'Invalid workbook engine format version',
      );
      assert(
        head.engineFormatVersion === 1,
        'UNSUPPORTED_VERSION',
        `Workbook engine format ${head.engineFormatVersion} is unsupported; use a compatible version`,
      );
    }
    assert(
      Object.keys(head).every((key) =>
        [
          'engineFormatVersion',
          'directory',
          'meta',
          'revision',
          'past',
          'future',
          'historyBytes',
          'budget',
        ].includes(key),
      ),
      'CORRUPT_STORAGE',
      'Unknown workbook head fields; open with a compatible version or restore a backup',
    );
    const owned = (value: unknown) => typeof value === 'string' && value.startsWith(`${this.id}/`);
    assert(
      owned(head.meta) &&
        Number.isSafeInteger(head.revision) &&
        head.revision >= 0 &&
        Number.isSafeInteger(head.budget) &&
        head.budget > 0 &&
        Number.isSafeInteger(head.historyBytes) &&
        head.historyBytes >= 0 &&
        head.historyBytes <= 32 * 1024 * 1024 &&
        Array.isArray(head.past) &&
        Array.isArray(head.future) &&
        head.past.length + head.future.length <= 100 &&
        [...head.past, ...head.future].every(owned) &&
        new Set([...head.past, ...head.future]).size === head.past.length + head.future.length,
      'CORRUPT_STORAGE',
      'Invalid workbook head history',
    );
    assert(
      head.directory &&
        typeof head.directory === 'object' &&
        !Array.isArray(head.directory) &&
        !Object.hasOwn(head, 'blocks') &&
        Object.entries(head.directory).every(
          ([bucket, id]) =>
            /^(0|[1-9]\d*)$/.test(bucket) &&
            Number(bucket) < 256 &&
            typeof id === 'string' &&
            id.startsWith(`${this.id}/`),
        ),
      'CORRUPT_STORAGE',
      'Unsupported workbook database format',
    );
    const blocks: Record<string, Pointer> = {};
    this.directoryEntries.clear();
    for (const [bucket, id] of Object.entries(head.directory)) {
      const entries = await this.db.get<Record<string, Pointer>>('records', id);
      assert(
        entries && typeof entries === 'object' && !Array.isArray(entries),
        'CORRUPT_STORAGE',
        'Invalid block directory',
      );
      for (const [address, pointer] of Object.entries(entries)) {
        const match = /^([\w-]+)\/(0|[1-9]\d*)\/(0|[1-9]\d*)$/.exec(address);
        assert(
          match &&
            Number(match[2]) <= Math.floor(0xffffffff / BLOCK_ROWS) &&
            Number(match[3]) <= Math.floor(0xffffffff / BLOCK_COLUMNS) &&
            this.bucket(address) === bucket &&
            !Object.hasOwn(blocks, address) &&
            this.validPointer(pointer),
          'CORRUPT_STORAGE',
          'Invalid block directory entry',
        );
      }
      this.directoryEntries.set(bucket, entries);
      Object.assign(blocks, entries);
    }
    return { ...head, blocks };
  }
  private async publish(next: Head, metadata: Metadata, ranges?: Selection[]) {
    this.check();
    const previousDirectory = this.head.directory ?? {};
    const directory = { ...previousDirectory },
      dirty = new Map<string, Record<string, Pointer>>();
    for (const address of this.changed) {
      const bucket = this.bucket(address);
      if (!dirty.has(bucket)) dirty.set(bucket, { ...this.directoryEntries.get(bucket) });
      const entries = dirty.get(bucket)!;
      if (next.blocks[address]) entries[address] = next.blocks[address];
      else delete entries[address];
    }
    const oldDirectoryIds: string[] = [];
    for (const [bucket, entries] of dirty) {
      if (directory[bucket]) oldDirectoryIds.push(directory[bucket]);
      if (Object.keys(entries).length)
        directory[bucket] = await this.allocate(entries, next.budget);
      else delete directory[bucket];
    }
    next.directory = directory;
    next.engineFormatVersion = 1;
    const bytes =
      this.usedBytes -
      recordBytes(this.id, this.storedHead(this.head)) +
      recordBytes(this.id, this.storedHead(next));
    assert(bytes <= next.budget, 'STORAGE_BUDGET', 'Application storage budget exceeded');
    this.check();
    await this.db.write('heads', [[this.id, this.storedHead(next)]]);
    try {
      const old = this.head;
      if ([...this.changed].some((a) => !!old.blocks[a] !== !!next.blocks[a]))
        this.exportAddresses = undefined;
      for (const id of Object.values(directory))
        if (!Object.values(previousDirectory).includes(id)) this.retain(id, 1);
      for (const id of oldDirectoryIds) this.retain(id, -1);
      for (const [bucket, entries] of dirty) {
        if (Object.keys(entries).length) this.directoryEntries.set(bucket, entries);
        else this.directoryEntries.delete(bucket);
      }
      const full = next.blocks !== old.blocks && !this.changed.size;
      if (full) {
        for (const p of Object.values(next.blocks)) this.retain(p.id, 1);
        for (const p of Object.values(old.blocks)) this.retain(p.id, -1);
      } else
        for (const address of this.changed) {
          const a = old.blocks[address],
            b = next.blocks[address];
          if (a?.id === b?.id) continue;
          if (b) this.retain(b.id, 1);
          if (a) this.retain(a.id, -1);
        }
      if (next.meta !== old.meta) {
        this.retain(next.meta, 1);
        this.retain(old.meta, -1);
      }
      const oldHistory = new Set([...old.past, ...old.future]),
        newHistory = new Set([...next.past, ...next.future]);
      for (const id of newHistory)
        if (!oldHistory.has(id)) {
          const h = this.histories.get(id) ?? (await this.readHistory(id));
          assert(h, 'CORRUPT_STORAGE', 'Missing new history');
          this.histories.set(id, h);
          this.retainHistory(id, h, 1);
        }
      for (const id of oldHistory)
        if (!newHistory.has(id)) {
          const h = this.histories.get(id)!;
          this.retainHistory(id, h, -1);
          this.histories.delete(id);
        }
      const priorSheets = new Map(this.metadata.snapshot.sheets.map((sh) => [sh.id, sh]));
      this.head = next;
      this.metadata = metadata;
      if (old.meta !== next.meta) {
        for (const sh of metadata.snapshot.sheets) {
          const prior = priorSheets.get(sh.id);
          // Compare identity order only when metadata changes; scalar commits skip this work.
          if (
            !prior ||
            prior.merges !== sh.merges ||
            prior.rowOrder.length !== sh.rowOrder.length ||
            prior.columnOrder.length !== sh.columnOrder.length ||
            prior.rowOrder.some((id, i) => id !== sh.rowOrder[i]) ||
            prior.columnOrder.some((id, i) => id !== sh.columnOrder[i])
          )
            delete this.used[sh.id];
        }
      }
      for (const a of this.changed)
        if (
          old.blocks[a]?.rowMask !== next.blocks[a]?.rowMask ||
          old.blocks[a]?.columnMask !== next.blocks[a]?.columnMask
        )
          delete this.used[a.split('/')[0]];
      this.usedBytes = bytes;
      if (ranges) this.calculator.invalidate(ranges);
      else this.calculator.clear();
      await this.cleanup(oldDirectoryIds);
    } catch {
      this.poisoned = true;
      throw new OpenSheetError('WORKER_FAILED', 'Commit may have succeeded; reopen the workbook');
    }
  }
  view(): EngineView {
    const used = this.used;
    for (const s of this.metadata.snapshot.sheets) {
      if (Object.hasOwn(used, s.id)) continue;
      let endRow = 1,
        endColumn = 1;
      const rows = this.axis(s.rowOrder),
        cols = this.axis(s.columnOrder);
      for (const [address, p] of Object.entries(this.head.blocks))
        if (address.startsWith(`${s.id}/`)) {
          const [, br, bc] = address.split('/'),
            mask = BigInt('0x' + (p.rowMask ?? '0'));
          for (let i = 0; i < 64; i++)
            if (mask & (1n << BigInt(i)))
              endRow = Math.max(endRow, (rows.get(+br * 64 + i) ?? -1) + 1);
          for (let i = 0; i < 32; i++)
            if ((p.columnMask ?? 0) & (1 << i))
              endColumn = Math.max(endColumn, (cols.get(+bc * 32 + i) ?? -1) + 1);
        }
      for (const m of s.merges) {
        endRow = Math.max(endRow, m.endRow);
        endColumn = Math.max(endColumn, m.endColumn);
      }
      used[s.id] = { startRow: 0, startColumn: 0, endRow, endColumn };
    }
    return {
      snapshot: { ...this.metadata.snapshot, revision: this.head.revision },
      metadataId: this.head.meta,
      revision: this.head.revision,
      importRevision: this.metadata.importRevision,
      filtered: this.metadata.filtered,
      used,
      canUndo: !!this.head.past.length,
      canRedo: !!this.head.future.length,
      bytes: this.usedBytes,
      cacheBytes: this.cache.bytes + this.calculator.cacheBytes + this.seen.bytes,
      budget: this.head.budget,
    };
  }
  private async block(address: string, head = this.head, timing?: ReadTiming): Promise<ColumnPage> {
    const p = head.blocks[address];
    if (!p) {
      if (timing) timing.absentPages++;
      return decodeBlock(encodeBlock({}, [], [], {}));
    }
    const cached = this.cache.get(p.id);
    if (cached) {
      if (timing) timing.cacheHits++;
      return cached;
    }
    if (timing) timing.cacheMisses++;
    let started = timing ? performance.now() : 0;
    const encoded = await this.db.get<EncodedBlock>('records', p.id);
    if (timing) {
      timing.storageMs += performance.now() - started;
      started = performance.now();
    }
    assert(encoded, 'CORRUPT_STORAGE', 'Data block is missing');
    const inflated = await this.inflate(encoded);
    if (timing) {
      const now = performance.now();
      timing.inflateMs += now - started;
      started = now;
    }
    const block = decodeBlock(inflated);
    if (timing) {
      const duration = performance.now() - started;
      timing.pageDecodeMs += duration;
      timing.decodeMs = timing.inflateMs + timing.pageDecodeMs;
    }
    this.cache.set(p.id, block, block.bytes);
    return block;
  }
  async cell(sheet: string, row: number, column: number): Promise<CellRecord | undefined> {
    return this.stagedCell(sheet, row, column, this.head, this.metadata);
  }
  private async stagedCell(
    sheet: string,
    row: number,
    column: number,
    head: Head,
    metadata: Metadata,
    local?: StagedReadPage,
  ): Promise<CellRecord | undefined> {
    this.check();
    const sh = metadata.snapshot.sheets.find((s) => s.id === sheet);
    if (
      !sh ||
      row < 0 ||
      column < 0 ||
      row >= sh.rowOrder.length ||
      column >= sh.columnOrder.length
    )
      return { rowId: 0, columnId: 0, input: { type: 'error', code: '#REF!' } };
    const address = blockAddress(sheet, sh.rowOrder[row], sh.columnOrder[column]);
    const page = await this.stagedPage(address, head, local);
    return page.cell((sh.rowOrder[row] % 64) * 32 + (sh.columnOrder[column] % 32));
  }
  private async stagedPage(address: string, head: Head, local?: StagedReadPage) {
    let page = local?.address === address ? local.page : undefined;
    if (!page) {
      page = await this.block(address, head);
      if (local) {
        local.address = page.bytes <= CONDITIONAL_READ_PAGE_BUDGET ? address : undefined;
        local.page = page.bytes <= CONDITIONAL_READ_PAGE_BUDGET ? page : undefined;
      }
    }
    return page;
  }
  /** Internal bridge; caller must keep head and metadata immutable until validation returns. */
  private async validateStagedCells(
    head: Head,
    metadata: Metadata,
    sheetId: string,
    rules: readonly ValidationEntry[],
    positions: readonly { row: number; column: number }[],
    signal?: AbortSignal,
  ) {
    this.check();
    signal?.throwIfAborted();
    const sheet = metadata.snapshot.sheets.find((candidate) => candidate.id === sheetId);
    assert(sheet, 'INVALID_ARGUMENT', 'Unknown validation sheet');
    const calculator = new AsyncCalculation(
      (id, row, column) => this.stagedCell(id, row, column, head, metadata),
      (name) =>
        metadata.snapshot.sheets.find(
          (candidate) => candidate.name.toLowerCase() === name.toLowerCase(),
        )?.id,
      this.calculator.cacheBudget,
      undefined,
      () => metadata.snapshot.dateSystem,
    );
    const reader = calculator.validationSession(signal).forSheet(sheetId);
    const report = await checkStagedValidation(
      rules,
      positions,
      reader.read,
      signal,
      reader.evaluateFormula,
      { rows: sheet.rowOrder.length, columns: sheet.columnOrder.length },
    );
    this.check();
    signal?.throwIfAborted();
    return report;
  }
  /** Internal bridge for directly edited ranges; formula dependents require separate enumeration. */
  private async validateStagedRanges(
    head: Head,
    metadata: Metadata,
    sheetId: string,
    rules: readonly ValidationEntry[],
    ranges: readonly Rect[],
    signal?: AbortSignal,
  ) {
    this.check();
    const sheet = metadata.snapshot.sheets.find((candidate) => candidate.id === sheetId);
    assert(sheet, 'INVALID_ARGUMENT', 'Unknown validation sheet');
    const selected = await validationCandidates(
      rules,
      ranges,
      { rows: sheet.rowOrder.length, columns: sheet.columnOrder.length },
      signal,
    );
    this.check();
    return this.validateStagedCells(
      head,
      metadata,
      sheetId,
      selected.rules,
      selected.positions,
      signal,
    );
  }
  /** Internal bridge; caller owns an immutable head and metadata until completion. */
  private async conditionalStagedViewport(
    head: Head,
    metadata: Metadata,
    sheetId: string,
    rules: readonly ConditionalRule[],
    positions: readonly { row: number; column: number }[],
    signal?: AbortSignal,
  ): Promise<CellStyle[]> {
    this.check();
    signal?.throwIfAborted();
    const sheet = metadata.snapshot.sheets.find((candidate) => candidate.id === sheetId);
    assert(sheet, 'INVALID_ARGUMENT', 'Unknown conditional formatting sheet');
    const checked = validateConditionalRules(
      rules,
      sheet.rowOrder.length,
      sheet.columnOrder.length,
    );
    assert(
      Array.isArray(positions) &&
        positions.length <= 10_000 &&
        positions.length * checked.length <= 1_000_000,
      'LIMIT_EXCEEDED',
      'Conditional formatting viewport exceeds work budget',
    );
    const snapshot = positions.map(({ row, column }) => {
      assert(
        Number.isInteger(row) &&
          row >= 0 &&
          row < sheet.rowOrder.length &&
          Number.isInteger(column) &&
          column >= 0 &&
          column < sheet.columnOrder.length,
        'INVALID_ARGUMENT',
        'Conditional formatting cell is outside the sheet',
      );
      return { row, column };
    });
    let current: { row: number; column: number; cell: CellRecord | undefined } | undefined;
    // Retain only scalar inputs already read for this bounded viewport. Formula
    // cells are deliberately excluded: cycle errors depend on traversal context.
    const scalarMemo = checked.some((rule) => rule.condition.kind === 'formula')
      ? new Map<number, CellRecord | undefined>()
      : undefined;
    let reuseScalars = false;
    const calculator = new AsyncCalculation(
      (id, row, column) => {
        if (id === sheetId && current?.row === row && current.column === column)
          return Promise.resolve(current.cell);
        const address = row * LIMITS.columns + column;
        if (id === sheetId && reuseScalars && scalarMemo?.has(address))
          return Promise.resolve(scalarMemo.get(address));
        return this.stagedCell(id, row, column, head, metadata);
      },
      (name) =>
        metadata.snapshot.sheets.find(
          (candidate) => candidate.name.toLowerCase() === name.toLowerCase(),
        )?.id,
      this.calculator.cacheBudget,
      undefined,
      () => metadata.snapshot.dateSystem,
    );
    const reader = calculator.validationSession(signal).forSheet(sheetId);
    const cells: Array<{ row: number; column: number; value: CellValue; base: CellStyle }> = [];
    const viewportPage: StagedReadPage = {};
    let valueText = 0;
    for (const { row, column } of snapshot) {
      this.check();
      signal?.throwIfAborted();
      const cell = await abortable(
        this.stagedCell(sheetId, row, column, head, metadata, viewportPage),
        signal,
      );
      current = { row, column, cell };
      let value: CellValue;
      try {
        value = await abortable(reader.read(row, column), signal);
      } finally {
        current = undefined;
      }
      signal?.throwIfAborted();
      assert(
        !(value && typeof value === 'object' && value.error === '#LIMIT!'),
        'LIMIT_EXCEEDED',
        'Conditional formatting calculation exceeded its work budget',
      );
      if (typeof value === 'string') {
        valueText += value.length;
        assert(
          valueText <= 10_000_000,
          'LIMIT_EXCEEDED',
          'Conditional formatting viewport exceeds text work budget',
        );
      }
      if (scalarMemo && (!cell || cell.input.type !== 'formula'))
        scalarMemo.set(
          row * LIMITS.columns + column,
          cell ? { rowId: cell.rowId, columnId: cell.columnId, input: cell.input } : undefined,
        );
      cells.push({
        row,
        column,
        value,
        base: cell?.styleId ? (metadata.snapshot.styles[cell.styleId] ?? {}) : {},
      });
    }
    reuseScalars = true;
    const result = await conditionalViewportWithFormulas(
      checked,
      cells,
      reader.evaluateFormula,
      signal,
    );
    this.check();
    signal?.throwIfAborted();
    return result;
  }
  /** Find the next occupied-region boundary using stored inputs, not calculated display values. */
  async navigateData(
    sheetId: string,
    row: number,
    column: number,
    axis: 'row' | 'column',
    direction: -1 | 1,
  ) {
    validateNavigationArguments(sheetId, row, column, axis, direction);
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Unknown sheet');
    assert(
      row < sh.rowOrder.length && column < sh.columnOrder.length,
      'INVALID_ARGUMENT',
      'Invalid navigation position',
    );
    const head = this.head,
      revision = head.revision;
    const vertical = axis === 'row',
      start = vertical ? row : column;
    const order = vertical ? sh.rowOrder : sh.columnOrder;
    const axisMeta = vertical ? sh.rows : sh.columns;
    const filtered = vertical
      ? new Set(
          Object.hasOwn(this.metadata.filtered, sheetId) ? this.metadata.filtered[sheetId] : [],
        )
      : undefined;
    const intervals = sh.merges
      .filter((m) =>
        vertical
          ? column >= m.startColumn && column < m.endColumn
          : row >= m.startRow && row < m.endRow,
      )
      .map((m) => ({
        start: vertical ? m.startRow : m.startColumn,
        end: vertical ? m.endRow : m.endColumn,
        row: m.startRow,
        column: m.startColumn,
      }))
      .sort((a, b) => a.start - b.start);
    const mergeAt = (index: number) => {
      let lo = 0,
        hi = intervals.length;
      while (lo < hi) {
        const mid = (lo + hi) >>> 1;
        if (intervals[mid].start <= index) lo = mid + 1;
        else hi = mid;
      }
      const merge = intervals[lo - 1];
      return merge && index < merge.end ? merge : undefined;
    };
    const visible = (index: number) =>
      !filtered?.has(index) &&
      !axisMeta[order[index]]?.hidden &&
      axisMeta[order[index]]?.size !== 0;
    let cachedAddress = '',
      cachedPage: ColumnPage | undefined;
    const occupied = (index: number): boolean | Promise<boolean> => {
      const merge = mergeAt(index),
        r = merge?.row ?? (vertical ? index : row),
        c = merge?.column ?? (vertical ? column : index);
      const rowId = sh.rowOrder[r],
        columnId = sh.columnOrder[c],
        address = blockAddress(sheetId, rowId, columnId);
      if (!head.blocks[address]) return false;
      const isFilled = (page: ColumnPage) =>
        (page.cell((rowId % 64) * 32 + (columnId % 32))?.input.type ?? 'blank') !== 'blank';
      if (address === cachedAddress) return isFilled(cachedPage!);
      return this.block(address, head).then((page) => {
        cachedPage = page;
        cachedAddress = address;
        return isFilled(page);
      });
    };
    const first = occupied(start),
      initial = typeof first === 'boolean' ? first : await first,
      startMerge = mergeAt(start);
    this.check();
    let index = startMerge
      ? direction > 0
        ? startMerge.end
        : startMerge.start - 1
      : start + direction;
    let last = start,
      phase: 'first' | 'occupied' | 'search' = 'first',
      checked = 0;
    while (index >= 0 && index < order.length) {
      this.check();
      if (++checked % 1024 === 0) {
        await yieldToTask();
        this.check();
      }
      if (!visible(index)) {
        index += direction;
        continue;
      }
      const current = occupied(index);
      const filled = typeof current === 'boolean' ? current : await current;
      this.check();
      if (phase === 'first') phase = initial && filled ? 'occupied' : 'search';
      if (phase === 'occupied' && !filled) break;
      last = index;
      if (phase === 'search' && filled) break;
      index += direction;
    }
    return { revision, row: vertical ? last : row, column: vertical ? column : last };
  }
  async read(sheetId: string, range: Rect, values = true, render = true, timing?: ReadTiming) {
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Unknown sheet');
    checkBounds(sh, range);
    assert(
      (range.endRow - range.startRow) * (range.endColumn - range.startColumn) <= 16384,
      'LIMIT_EXCEEDED',
      'Read ranges in batches of at most 16384 cells',
    );
    const cells: Record<string, CellRecord> = {},
      calculated: Record<string, CellValue> = {},
      display: Record<string, string> = {};
    let loadedAddress: string | undefined;
    let loadedPage: ColumnPage | undefined;
    let calculate: ReturnType<AsyncCalculation['readSession']> | undefined;
    let visited = 0;
    for (let row = range.startRow; row < range.endRow; row++)
      for (let column = range.startColumn; column < range.endColumn; column++) {
        this.check();
        if (++visited % 1024 === 0) {
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
          this.check();
        }
        const rowId = sh.rowOrder[row],
          columnId = sh.columnOrder[column],
          k = key(rowId, columnId),
          address = blockAddress(sheetId, rowId, columnId);
        if (address !== loadedAddress) {
          loadedPage = await this.block(address, this.head, timing);
          loadedAddress = address;
        }
        const cell = loadedPage!.cell((rowId % 64) * 32 + (columnId % 32));
        if (cell) {
          cells[k] = cell;
          if (this.head.revision !== this.metadata.importRevision && cell.cached) {
            cells[k] = { ...cell };
            delete cells[k].cached;
          }
        }
        if (values) {
          const v =
            cell?.input.type === 'formula'
              ? await (calculate ??= this.calculator.readSession(16 * 1024))(sheetId, row, column)
              : cell
                ? scalarValue(cell.input)
                : null;
          calculated[k] = v;
          if (render)
            display[k] =
              v &&
              typeof v === 'object' &&
              cell?.cached &&
              this.head.revision === this.metadata.importRevision
                ? formatValue(
                    scalarValue(cell.cached),
                    cell.numberFormat,
                    this.metadata.snapshot.dateSystem,
                  ) + ' †'
                : formatValue(
                    v,
                    cell?.numberFormat ??
                      this.metadata.snapshot.styles[cell?.styleId ?? '']?.numberFormat,
                    this.metadata.snapshot.dateSystem,
                  );
        }
      }
    return { sheetId, range, cells, calculated, display, revision: this.head.revision };
  }
  async readCells(sheetId: string, range: Rect, fields: import('./workbook.js').CellField[]) {
    const values = fields.includes('value') || fields.includes('display');
    const data = await this.read(sheetId, range, values, fields.includes('display'));
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId)!;
    const cells: import('./workbook.js').CellData[] = [];
    for (let row = range.startRow; row < range.endRow; row++)
      for (let column = range.startColumn; column < range.endColumn; column++) {
        const k = key(sh.rowOrder[row], sh.columnOrder[column]),
          cell = data.cells[k];
        cells.push(this.selectCell(row, column, cell, fields, data.calculated[k], data.display[k]));
      }
    return { revision: this.head.revision, sheetId, range, cells };
  }
  private selectCell(
    row: number,
    column: number,
    cell: CellRecord | undefined,
    fields: import('./workbook.js').CellField[],
    value?: CellValue,
    display?: string,
  ) {
    const output: import('./workbook.js').CellData = { row, column };
    for (const field of fields) {
      if (field === 'input') output.input = cell?.input ?? { type: 'blank' };
      else if (field === 'value') output.value = value ?? null;
      else if (field === 'display') output.display = display ?? '';
      else if (field === 'style')
        output.style = cell?.styleId ? this.metadata.snapshot.styles[cell.styleId] : {};
      else if (field === 'numberFormat')
        output.numberFormat =
          cell?.numberFormat ?? this.metadata.snapshot.styles[cell?.styleId ?? '']?.numberFormat;
      else if (field === 'note') output.note = cell?.note;
      else if (field === 'link') output.link = cell?.link;
    }
    return output;
  }
  async scanCells(
    sheetId: string,
    options: {
      cursor?: import('./workbook.js').CellCursor;
      limit?: number;
      fields?: import('./workbook.js').CellField[];
    },
  ) {
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Unknown sheet');
    const revision = this.head.revision,
      limit = options.limit ?? 1024,
      fields = options.fields ?? ['input'];
    assert(
      Number.isInteger(limit) && limit > 0 && limit <= 2048,
      'INVALID_ARGUMENT',
      'Cell page limit must be 1..2048',
    );
    const cursor = options.cursor;
    if (cursor)
      assert(
        cursor.revision === revision && cursor.sheetId === sheetId,
        'REVISION_CONFLICT',
        'Cell cursor expired',
      );
    if (cursor)
      assert(
        Number.isInteger(cursor.offset) &&
          cursor.offset >= 0 &&
          !!this.head.blocks[cursor.address] &&
          cursor.address.startsWith(`${sheetId}/`),
        'INVALID_ARGUMENT',
        'Invalid cell cursor',
      );
    const address = cursor?.address ?? this.exportIndex(revision, undefined, 1, sheetId)[0];
    if (!address) return { revision, cells: [] };
    const page = await this.block(address),
      source = Object.values(page.cells).sort(
        (a, b) => a.rowId - b.rowId || a.columnId - b.columnId,
      ),
      start = cursor?.offset ?? 0;
    assert(start <= source.length, 'INVALID_ARGUMENT', 'Invalid cell cursor offset');
    const cells: import('./workbook.js').CellData[] = [];
    for (const cell of source.slice(start, start + limit)) {
      this.check();
      const row = this.axis(sh.rowOrder).get(cell.rowId)!,
        column = this.axis(sh.columnOrder).get(cell.columnId)!;
      const value =
        fields.includes('value') || fields.includes('display')
          ? await this.calculator.value(sheetId, row, column)
          : undefined;
      cells.push(
        this.selectCell(
          row,
          column,
          cell,
          fields,
          value,
          fields.includes('display')
            ? formatValue(
                value &&
                  typeof value === 'object' &&
                  cell.cached &&
                  this.head.revision === this.metadata.importRevision
                  ? scalarValue(cell.cached)
                  : (value ?? null),
                cell.numberFormat ??
                  this.metadata.snapshot.styles[cell.styleId ?? '']?.numberFormat,
                this.metadata.snapshot.dateSystem,
              ) +
                (value &&
                typeof value === 'object' &&
                cell.cached &&
                this.head.revision === this.metadata.importRevision
                  ? ' †'
                  : '')
            : undefined,
        ),
      );
    }
    const offset = start + cells.length,
      nextAddress =
        offset < source.length ? address : this.exportIndex(revision, address, 1, sheetId)[0];
    return {
      revision,
      cells,
      ...(nextAddress
        ? {
            cursor: {
              revision,
              sheetId,
              address: nextAddress,
              offset: nextAddress === address ? offset : 0,
            },
          }
        : {}),
    };
  }
  private async storeBlock(
    address: string,
    cells: Record<string, CellRecord>,
    head: Head,
    metadata: Metadata,
    warmCache = true,
  ) {
    const sheetId = address.split('/')[0],
      sh = metadata.snapshot.sheets.find((s) => s.id === sheetId)!;
    let index = this.indices.get(sh);
    if (!index) {
      index = {
        rows: this.axis(sh.rowOrder),
        cols: this.axis(sh.columnOrder),
      };
      this.indices.set(sh, index);
    }
    const { rows, cols } = index;
    this.changed.add(address);
    const previous = head.blocks[address];
    const discardPrevious = async () => {
      if (previous && this.allocated.has(previous.id) && !this.pins.has(previous.id)) {
        await this.db.write('records', [], [previous.id]);
        this.usedBytes -= this.sizes.get(previous.id) ?? 0;
        this.sizes.delete(previous.id);
        this.cache.delete(previous.id);
        this.allocated.delete(previous.id);
      }
    };
    const formulas = new Set<string>();
    const styles = new Set<string>();
    const visit = (node: AST) => {
      if (node.kind === 'ref') formulas.add(node.sheet ?? sh.name);
      else if (node.kind === 'range') {
        formulas.add(node.start.sheet ?? sh.name);
        formulas.add(node.end.sheet ?? sh.name);
      } else if (node.kind === 'unary') visit(node.value);
      else if (node.kind === 'binary') {
        visit(node.left);
        visit(node.right);
      } else if (node.kind === 'call') node.args.forEach(visit);
    };
    let retained = 0,
      rowLow = 0,
      rowHigh = 0,
      columnMask = 0,
      count = 0,
      endRow = 1,
      endColumn = 1;
    for (const [k, c] of Object.entries(cells)) {
      if (
        c.input.type === 'blank' &&
        c.styleId === undefined &&
        c.numberFormat === undefined &&
        c.link === undefined &&
        c.note === undefined
      ) {
        delete cells[k];
        continue;
      }
      retained++;
      const rowBit = c.rowId % 64;
      if (rowBit < 32) rowLow |= 1 << rowBit;
      else rowHigh |= 1 << (rowBit - 32);
      columnMask |= 1 << c.columnId % 32;
      if (c.styleId !== undefined) styles.add(c.styleId);
      if (c.input.type !== 'blank') count++;
      endRow = Math.max(endRow, rows.get(c.rowId)! + 1);
      endColumn = Math.max(endColumn, cols.get(c.columnId)! + 1);
      if (c.input.type === 'formula') {
        try {
          visit(parseFormula(c.input.expression));
        } catch {
          formulas.add('*');
        }
      }
    }
    if (!retained) {
      delete head.blocks[address];
      await discardPrevious();
      return;
    }
    const encoded = encodeBlock(cells, [], [], {});
    const id = await this.allocate(encoded, head.budget);
    this.compressionQueue.add(id);
    const low = (rowLow >>> 0).toString(16);
    head.blocks[address] = {
      id,
      bytes: this.sizes.get(id)!,
      count,
      endRow,
      endColumn,
      rowMask: rowHigh ? (rowHigh >>> 0).toString(16) + low.padStart(8, '0') : low,
      columnMask: columnMask >>> 0,
      ...(formulas.size ? { formulas: [...formulas] } : {}),
      ...(styles.size ? { styles: [...styles] } : {}),
    };
    if (warmCache) {
      const page = decodeBlock(encoded);
      this.cache.set(id, page, page.bytes);
    }
    await discardPrevious();
  }
  async replace(snapshot: WorkbookSnapshot, budget = this.head.budget) {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.cancelled = false;
    this.allocations = [];
    this.changed.clear();
    this.allocated.clear();
    this.pins.clear();
    // Validate metadata separately; each cell is validated by the importer in bounded batches.
    const metadata: Metadata = {
      importRevision: snapshot.revision,
      snapshot: validateSnapshot({
        ...snapshot,
        sheets: snapshot.sheets.map((s) => ({ ...s, cells: {} })),
      }),
      filtered: {},
    };
    const next: Head = {
      meta: '',
      blocks: {},
      revision: snapshot.revision,
      past: [],
      future: [],
      historyBytes: 0,
      budget,
    };
    const old = await this.liveIds();
    for (const address of Object.keys(this.head.blocks)) this.changed.add(address);
    try {
      for (const sh of snapshot.sheets) {
        const rows = new Map(sh.rowOrder.map((id, i) => [id, i])),
          cols = new Map(sh.columnOrder.map((id, i) => [id, i]));
        const grouped = new Map<string, Record<string, CellRecord>>();
        // Snapshot import is an explicit full-memory compatibility boundary; generation uses generate().
        for (const [k, c] of Object.entries(sh.cells)) {
          const r = rows.get(c.rowId),
            col = cols.get(c.columnId);
          assert(
            r !== undefined && col !== undefined && k === key(c.rowId, c.columnId),
            'INVALID_ARGUMENT',
            'Invalid cell identity',
          );
          const a = blockAddress(sh.id, c.rowId, c.columnId);
          if (!grouped.has(a)) grouped.set(a, {});
          grouped.get(a)![k] = c;
        }
        let n = 0;
        for (const [address, cells] of grouped) {
          validateSnapshot({
            ...metadata.snapshot,
            sheets: metadata.snapshot.sheets.map((s) => (s.id === sh.id ? { ...s, cells } : s)),
          });
          await this.storeBlock(address, cells, next, metadata);
          this.progress({ phase: 'import', completed: ++n, total: grouped.size });
        }
      }
      await this.filters(next, metadata);
      next.meta = await this.allocate(encodeMetadata(metadata), budget);
      await this.publish(next, metadata);
      await this.cleanup([...old, ...this.allocations]);
      return this.view();
    } catch (error) {
      await this.cleanup(this.allocations);
      throw error;
    }
  }
  async importJSON(input: WorkbookFile | JSONSource, budget = this.head.budget) {
    return this.importStaged(input, budget);
  }
  /** Internal migration path. Public native import continues to reject schema 4. */
  async importRuleFreeValidationBackup(input: JSONSource, budget = this.head.budget) {
    assert(
      typeof input === 'string' || input instanceof Blob || input instanceof ReadableStream,
      'INVALID_ARGUMENT',
      'Migration requires a JSON source',
    );
    return this.importStaged(input, budget, downgradeValidationMetadata);
  }
  private async importStaged(
    input: WorkbookFile | JSONSource,
    budget: number,
    convertMetadata?: (input: WorkbookFile) => WorkbookFile,
  ) {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.check();
    this.allocations = [];
    this.changed.clear();
    this.allocated.clear();
    this.pins.clear();
    const old = await this.liveIds();
    this.check();
    for (const address of Object.keys(this.head.blocks)) this.changed.add(address);
    let file: WorkbookFile | undefined,
      binarySnapshot: WorkbookSnapshot | undefined,
      staged = 0;
    const scratch = `${this.id}/import`;
    let scratchWriteMs = 0;
    const sourceStageStart = performance.now();
    const reader = new AbortController();
    this.importRead = reader;
    const stage = async (value: unknown, suffix = '') => {
      this.check();
      const record: [IDBValidKey, unknown] = [
        `${scratch}/${String(staged++).padStart(10, '0')}${suffix}`,
        value,
      ];
      this.stagingBytes += recordBytes(record[0], record[1]);
      assert(
        this.usedBytes + this.stagingBytes <= budget,
        'STORAGE_BUDGET',
        'Import staging exceeds storage budget',
      );
      const writeStart = performance.now();
      await this.db.write('scratch', [record], [], reader.signal);
      scratchWriteMs += performance.now() - writeStart;
      this.progress({ phase: 'import-parse', completed: staged });
    };
    try {
      if (input && typeof input === 'object' && 'binary' in input) {
        for await (const item of readBinary(
          (input as unknown as { binary: Blob | Uint8Array | ReadableStream<Uint8Array> }).binary,
          { signal: reader.signal },
        )) {
          if ('metadata' in item) {
            file = item.metadata;
            binarySnapshot = item.snapshot;
          } else {
            await stage(item.bytes, `/${item.sheet}/${item.row}/${item.column}`);
          }
        }
      } else if (
        typeof input === 'string' ||
        input instanceof Blob ||
        input instanceof ReadableStream
      ) {
        for await (const item of parseWorkbookJSON(input, { signal: reader.signal })) {
          if ('block' in item) await stage(item);
          else file = convertMetadata ? convertMetadata(item.metadata) : item.metadata;
        }
      } else {
        assert(input && Array.isArray(input.sheets), 'INVALID_ARGUMENT', 'Invalid workbook file');
        file = { ...input, sheets: input.sheets.map((sh) => ({ ...sh, blocks: [] })) };
        for (let i = 0; i < input.sheets.length; i++)
          for (const block of input.sheets[i].blocks) await stage({ sheet: i, block });
      }
      this.progress({
        phase: 'import-source-stage',
        completed: staged,
        elapsedMs: performance.now() - sourceStageStart,
      });
      this.progress({
        phase: 'import-scratch-write',
        completed: staged,
        elapsedMs: scratchWriteMs,
      });
      this.check();
      assert(file, 'INVALID_ARGUMENT', 'Missing workbook metadata');
      const snapshot = binarySnapshot ?? decodeFileMetadata(file);
      snapshot.workbookId = this.id;
      // Import is a logical change even if the source revision is older.
      snapshot.revision = this.head.meta ? this.head.revision + 1 : file.revision;
      const metadata: Metadata = { snapshot, importRevision: snapshot.revision, filtered: {} };
      const next: Head = {
        meta: '',
        blocks: {},
        revision: snapshot.revision,
        past: [],
        future: [],
        historyBytes: 0,
        budget,
      };
      const styles = Object.keys(snapshot.styles),
        styleSet = binarySnapshot ? new Set(styles) : undefined,
        rows = snapshot.sheets.map((sh) => new Set(sh.rowOrder)),
        columns = snapshot.sheets.map((sh) => new Set(sh.columnOrder));
      // Filters may immediately read staged pages (including cross-sheet references).
      // Otherwise let the first actual read populate the cache.
      const warmCache = snapshot.sheets.some((sheet) => !!sheet.filter);
      let completed = 0;
      const seen = new Set<string>();
      const times = {
        'import-scratch-read': 0,
        'import-decode': 0,
        'import-pages': 0,
        'import-scratch-delete': 0,
      };
      let previous = performance.now();
      for await (const batch of this.db.scan<
        | Uint8Array
        | {
            sheet: number;
            block: import('./file-codec.js').FileBlock;
          }
      >('scratch', prefixRange(scratch), 16, 1024 * 1024)) {
        const received = performance.now();
        times['import-scratch-read'] += received - previous;
        for (const [id, item] of batch) {
          const decoding = performance.now();
          this.check();
          const binary = item instanceof Uint8Array;
          const [sheet, row, column] = binary
            ? String(id).split('/').slice(-3).map(Number)
            : [item.sheet, item.block.row, item.block.column];
          assert(
            Number.isInteger(sheet) && !!snapshot.sheets[sheet] && binary === !!binarySnapshot,
            'INVALID_ARGUMENT',
            'Unknown imported sheet',
          );
          const sh = snapshot.sheets[sheet],
            address = `${sh.id}/${row}/${column}`;
          assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate imported block');
          seen.add(address);
          const cells = binary
            ? decodeVerifiedBinaryPage(item, row, column, rows[sheet], columns[sheet], styleSet!)
            : decodeFileCells(item.block, styles, rows[sheet], columns[sheet]);
          const decoded = performance.now();
          times['import-decode'] += decoded - decoding;
          await this.storeBlock(address, cells, next, metadata, warmCache);
          const stored = performance.now();
          times['import-pages'] += stored - decoded;
          await this.db.write('scratch', [], [id]);
          times['import-scratch-delete'] += performance.now() - stored;
          this.stagingBytes -= recordBytes(id, item);
          this.progress({ phase: 'import-store', completed: ++completed, total: staged });
        }
        previous = performance.now();
      }
      times['import-scratch-read'] += performance.now() - previous;
      for (const [phase, elapsedMs] of Object.entries(times))
        this.progress({ phase, elapsedMs, completed, total: staged });
      await this.filters(next, metadata);
      next.meta = await this.allocate(encodeMetadata(metadata), budget);
      await this.publish(next, metadata);
      await this.cleanup([...old, ...this.allocations]);
      return this.view();
    } catch (error) {
      await this.cleanup(this.allocations);
      throw error;
    } finally {
      this.importRead = undefined;
      try {
        await this.db.write('scratch', [], [prefixRange(scratch)]);
        this.stagingBytes = 0;
      } catch {
        // Keep failed staging charged. Reopen reclaims it before another import
        // can reuse these keys. Preserve the original operation result.
        this.poisoned = true;
      }
    }
  }
  importBinary(input: Blob | Uint8Array | ReadableStream<Uint8Array>) {
    return this.importJSON({ binary: input } as unknown as WorkbookFile);
  }
  async generate(count: number, budget: number) {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    assert(
      Number.isSafeInteger(count) && count > 0 && count <= 10_000_000,
      'INVALID_ARGUMENT',
      'Count exceeds the 100-column dimension boundary',
    );
    this.cancelled = false;
    this.allocations = [];
    this.changed.clear();
    this.allocated.clear();
    this.pins.clear();
    const metadata: Metadata = structuredClone(this.metadata);
    const sh = metadata.snapshot.sheets[0];
    sh.rowOrder = Array.from({ length: Math.ceil(count / 100) }, (_, i) => i);
    sh.columnOrder = Array.from({ length: 100 }, (_, i) => i);
    sh.nextRowId = sh.rowOrder.length;
    sh.nextColumnId = sh.columnOrder.length;
    sh.cells = {};
    sh.rows = {};
    sh.columns = {};
    sh.merges = [];
    sh.freeze = { rows: 0, columns: 0 };
    sh.filter = null;
    metadata.snapshot.sheets = [sh];
    metadata.snapshot.sheetOrder = [sh.id];
    metadata.snapshot.styles = {};
    metadata.filtered = {};
    const next: Head = {
      meta: '',
      blocks: {},
      revision: this.head.revision + 1,
      past: [],
      future: [],
      historyBytes: 0,
      budget,
    };
    const old = await this.liveIds();
    for (const address of Object.keys(this.head.blocks)) this.changed.add(address);
    try {
      for (let r = 0; r < sh.rowOrder.length; r += BLOCK_ROWS)
        for (let c = 0; c < 100; c += BLOCK_COLUMNS) {
          const cells: Record<string, CellRecord> = {};
          for (let row = r; row < Math.min(r + BLOCK_ROWS, sh.rowOrder.length); row++)
            for (let column = c; column < Math.min(c + BLOCK_COLUMNS, 100); column++) {
              const n = row * 100 + column;
              if (n >= count) break;
              const v = fixtureValue(n);
              cells[key(sh.rowOrder[row], sh.columnOrder[column])] = {
                rowId: sh.rowOrder[row],
                columnId: sh.columnOrder[column],
                input:
                  typeof v === 'number'
                    ? { type: 'number', value: v }
                    : typeof v === 'boolean'
                      ? { type: 'boolean', value: v }
                      : { type: 'string', value: v },
              };
            }
          await this.storeBlock(blockAddress(sh.id, r, c), cells, next, metadata);
          this.progress({
            phase: 'generate',
            completed: Math.min(count, (r + BLOCK_ROWS) * 100),
            total: count,
          });
        }
      next.meta = await this.allocate(encodeMetadata(metadata), budget);
      await this.publish(next, metadata);
      await this.cleanup([...old, ...this.allocations]);
      return this.view();
    } catch (error) {
      await this.cleanup(this.allocations);
      throw error;
    }
  }
  private async filters(head: Head, metadata: Metadata) {
    metadata.filtered = {};
    if (!metadata.snapshot.sheets.some((sheet) => sheet.filter)) return;
    this.calculator.clear();
    const calculator = new AsyncCalculation(
      (sheet, row, column) => this.stagedCell(sheet, row, column, head, metadata),
      (name) =>
        metadata.snapshot.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())?.id,
      this.calculator.cacheBudget,
      undefined,
      () => metadata.snapshot.dateSystem,
    );
    for (const sh of metadata.snapshot.sheets)
      if (sh.filter) {
        const hidden: number[] = [];
        const filter = sh.filter;
        const query = 'column' in filter ? filter.query.toLowerCase() : undefined;
        const criteria =
          'criteria' in filter
            ? filter.criteria.map((criterion) => ({
                criterion,
                foldedQuery:
                  'value' in criterion && typeof criterion.value === 'string'
                    ? criterion.value.toLowerCase()
                    : undefined,
              }))
            : [];
        try {
          const headerRows = 'column' in filter ? 1 : filter.headerRows;
          for (let row = headerRows; row < sh.rowOrder.length; row++) {
            this.check();
            let matches: boolean;
            if ('column' in filter) {
              const c = await this.stagedCell(sh.id, row, filter.column, head, metadata);
              const value = await calculator.value(sh.id, row, filter.column);
              assert(
                !(value && typeof value === 'object' && value.error === '#LIMIT!'),
                'LIMIT_EXCEEDED',
                'Filter calculation exceeded its work budget',
              );
              matches = formatValue(
                value,
                c?.numberFormat ?? metadata.snapshot.styles[c?.styleId ?? '']?.numberFormat,
                metadata.snapshot.dateSystem,
              )
                .toLowerCase()
                .includes(query!);
            } else {
              matches = filter.mode === 'all';
              const values = new Map<number, CellValue>();
              for (const { criterion, foldedQuery } of criteria) {
                if (!values.has(criterion.column)) {
                  const value = await calculator.value(sh.id, row, criterion.column);
                  assert(
                    !(value && typeof value === 'object' && value.error === '#LIMIT!'),
                    'LIMIT_EXCEEDED',
                    'Filter calculation exceeded its work budget',
                  );
                  values.set(criterion.column, value);
                }
                const match = matchesFilter(values.get(criterion.column)!, criterion, foldedQuery);
                if (filter.mode === 'all' ? !match : match) {
                  matches = match;
                  break;
                }
              }
            }
            if (!matches) hidden.push(row);
          }
        } finally {
          calculator.clear();
        }
        metadata.filtered[sh.id] = hidden;
      }
  }
  private addresses(sh: SheetSnapshot, range: Rect) {
    const rows = new Set(
      sh.rowOrder.slice(range.startRow, range.endRow).map((id) => Math.floor(id / BLOCK_ROWS)),
    );
    const columns = new Set(
      sh.columnOrder
        .slice(range.startColumn, range.endColumn)
        .map((id) => Math.floor(id / BLOCK_COLUMNS)),
    );
    const out: string[] = [];
    for (const r of rows) for (const c of columns) out.push(`${sh.id}/${r}/${c}`);
    return out;
  }
  private async applyLocal(command: Command, head: Head, metadata: Metadata) {
    const p = command.payload as any,
      sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId);
    let initialCellCount = 0;
    let initialCellBytes = 0;
    if (command.type === 'core.sheet.add') {
      for (const position in p.sheet.cells) {
        if (!Object.hasOwn(p.sheet.cells, position)) continue;
        initialCellCount++;
        initialCellBytes += encodedBytes(p.sheet.cells[position]) * 4;
        assert(
          initialCellCount * 512 <= 32 * 1024 * 1024 && initialCellBytes <= 32 * 1024 * 1024,
          'LIMIT_EXCEEDED',
          'Initial sheet cells exceed the 32 MiB command budget; import as a new workbook',
        );
        if (initialCellCount % 1024 === 0) {
          await yieldToTask();
          this.check();
        }
      }
    }
    const traversalCells =
      (p.cells?.length ?? 0) +
      initialCellCount +
      [p.range, p.target].reduce(
        (total, range) =>
          total +
          (range ? (range.endRow - range.startRow) * (range.endColumn - range.startColumn) : 0),
        0,
      );
    let expectedTextBytes = 0;
    if (command.type === 'core.cells.set') {
      for (const item of p.cells) {
        const input = item.expected?.input;
        if (input?.type === 'string') expectedTextBytes += input.value.length * 2;
        else if (input?.type === 'formula') expectedTextBytes += input.expression.length * 2;
        else if (input?.type === 'error') expectedTextBytes += input.code.length * 2;
      }
    }
    assert(
      traversalCells * 512 + expectedTextBytes <= 32 * 1024 * 1024,
      'LIMIT_EXCEEDED',
      'Command exceeds the 32 MiB traversal planning budget; use a streamed operation',
    );
    const addresses = new Set<string>();
    if (sh) {
      if (p.cells)
        for (const c of p.cells)
          addresses.add(blockAddress(sh.id, sh.rowOrder[c.row], sh.columnOrder[c.column]));
      for (const range of [p.range, p.target])
        if (range) {
          checkBounds(sh, range);
          for (const a of this.addresses(sh, range)) addresses.add(a);
        }
    }
    let bytes = 0;
    const snapshot: WorkbookSnapshot = {
      ...metadata.snapshot,
      sheetOrder: [...metadata.snapshot.sheetOrder],
      styles: changesStyles(command) ? { ...metadata.snapshot.styles } : metadata.snapshot.styles,
      sheets: metadata.snapshot.sheets.map((s) => ({
        ...s,
        cells: {},
        rows:
          command.type === 'core.axis.meta' ||
          command.type === 'core.axis.visibility' ||
          command.type === 'core.axis.visibilityAll' ||
          command.type === 'core.axis.resize'
            ? { ...s.rows }
            : s.rows,
        columns:
          command.type === 'core.axis.meta' ||
          command.type === 'core.axis.visibility' ||
          command.type === 'core.axis.visibilityAll' ||
          command.type === 'core.axis.resize'
            ? { ...s.columns }
            : s.columns,
        merges: command.type === 'core.cells.merge' ? s.merges.map((m) => ({ ...m })) : s.merges,
        freeze: s.freeze,
        filter: s.filter,
      })),
    };
    const target = snapshot.sheets.find((s) => s.id === p.sheetId);
    // Merging changes metadata even when every covered cell is already blank.
    // Keep untouched page pointers rather than rewriting those pages.
    const mergeDirty = command.type === 'core.cells.merge' ? new Set<string>() : undefined;
    if (mergeDirty) assert(target, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    const mergeRows =
      mergeDirty && new Set(target!.rowOrder.slice(p.range.startRow, p.range.endRow));
    const mergeColumns =
      mergeDirty && new Set(target!.columnOrder.slice(p.range.startColumn, p.range.endColumn));
    const mergeAnchorRow = mergeDirty && target!.rowOrder[p.range.startRow];
    const mergeAnchorColumn = mergeDirty && target!.columnOrder[p.range.startColumn];
    for (const address of addresses) {
      const block = await this.block(address, head);
      bytes += encodedBytes(block) * 4;
      assert(
        bytes <= 32 * 1024 * 1024,
        'LIMIT_EXCEEDED',
        'Command working set exceeds 32 MiB; use a streamed operation',
      );
      if (mergeDirty)
        for (const cell of Object.values(block.cells))
          if (
            mergeRows!.has(cell.rowId) &&
            mergeColumns!.has(cell.columnId) &&
            (cell.rowId !== mergeAnchorRow || cell.columnId !== mergeAnchorColumn) &&
            (cell.input.type !== 'blank' || cell.cached !== undefined)
          ) {
            mergeDirty.add(address);
            break;
          }
      Object.assign(target!.cells, structuredClone(block.cells));
    }
    reduceCommand(snapshot, command);
    if (command.type === 'core.sheet.add') {
      const index = snapshot.sheets.findIndex((sheet) => sheet.id === p.sheet.id);
      const added = snapshot.sheets[index];
      const initialPages = new Map<string, Record<string, CellRecord>>();
      let grouped = 0;
      for (const [position, cell] of Object.entries(added.cells)) {
        const address = blockAddress(added.id, cell.rowId, cell.columnId);
        let page = initialPages.get(address);
        if (!page) initialPages.set(address, (page = {}));
        page[position] = cell;
        if (++grouped % 1024 === 0) {
          await yieldToTask();
          this.check();
        }
      }
      // Do not clear the sheet object retained in the returned command payload.
      snapshot.sheets[index] = { ...added, cells: {} };
      metadata.snapshot = snapshot;
      for (const [address, cells] of initialPages) {
        this.check();
        await this.storeBlock(address, cells, head, metadata);
      }
    }
    const pages = new Map<string, Record<string, CellRecord>>();
    if (target && (!mergeDirty || mergeDirty.size)) {
      for (const cell of Object.values(target.cells)) {
        const address = blockAddress(target.id, cell.rowId, cell.columnId);
        if (mergeDirty && !mergeDirty.has(address)) continue;
        let page = pages.get(address);
        if (!page) pages.set(address, (page = {}));
        page[key(cell.rowId, cell.columnId)] = cell;
      }
    }
    for (const address of addresses) {
      if (mergeDirty && !mergeDirty.has(address)) continue;
      await this.storeBlock(address, pages.get(address) ?? {}, head, metadata);
    }
    for (const s of snapshot.sheets) s.cells = {};
    metadata.snapshot = snapshot;
  }
  async execute(inputs: Array<Command | CommandEnvelope>, label = 'edit') {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.cancelled = false;
    this.allocations = [];
    this.changed.clear();
    this.allocated.clear();
    this.pins.clear();
    const envelope =
      inputs.length === 1 && 'protocolVersion' in inputs[0]
        ? (inputs[0] as CommandEnvelope)
        : undefined;
    let signature = '';
    if (envelope) {
      assert(
        envelope.protocolVersion === 1 &&
          envelope.workbookId === this.id &&
          typeof envelope.commandId === 'string' &&
          envelope.commandId.length > 0 &&
          envelope.commandId.length <= 100,
        'INVALID_ARGUMENT',
        'Invalid command envelope',
      );
      signature = JSON.stringify(envelope);
      const old = this.seen.get(envelope.commandId);
      if (old) {
        assert(old.signature === signature, 'INVALID_ARGUMENT', 'Command id reused');
        return { view: this.view(), commit: old.commit, duplicate: true };
      }
      assert(envelope.baseRevision === this.head.revision, 'REVISION_CONFLICT', 'Workbook changed');
    }
    const commands = inputs.map(validateCommand),
      old = this.head,
      metadata: Metadata = {
        ...this.metadata,
        snapshot: {
          ...this.metadata.snapshot,
          sheetOrder: [...this.metadata.snapshot.sheetOrder],
          styles: this.metadata.snapshot.styles,
          sheets: this.metadata.snapshot.sheets.map((sh) => ({
            ...sh,
            cells: {},
            rows: sh.rows,
            columns: sh.columns,
            merges: sh.merges,
            freeze: sh.freeze,
            filter: sh.filter,
          })),
        },
      },
      transaction = overlay(old.blocks),
      next: Head = {
        ...old,
        blocks: transaction.record,
        past: [...old.past],
        future: [...old.future],
      };
    // Resolve a standalone visibility command against the queued Worker revision.
    // A no-op must not create history or discard redo, including for envelopes.
    if (
      commands.length === 1 &&
      (commands[0].type === 'core.axis.visibility' ||
        commands[0].type === 'core.axis.visibilityAll')
    ) {
      const p = commands[0].payload as {
        sheetId: string;
        axis: 'row' | 'column';
        index: number;
        count: number;
        hidden: boolean;
      };
      const sheet = metadata.snapshot.sheets.find((item) => item.id === p.sheetId);
      assert(sheet, 'INVALID_ARGUMENT', 'Sheet no longer exists');
      const order = p.axis === 'row' ? sheet.rowOrder : sheet.columnOrder;
      const axisMetadata = p.axis === 'row' ? sheet.rows : sheet.columns;
      const start = commands[0].type === 'core.axis.visibilityAll' ? 0 : p.index;
      const end = commands[0].type === 'core.axis.visibilityAll' ? order.length : p.index + p.count;
      assert(end <= order.length, 'INVALID_RANGE', 'Unknown row/column');
      let changes = false;
      for (let i = start; i < end && !changes; i++)
        changes = !!axisMetadata[order[i]]?.hidden !== p.hidden;
      if (!changes) {
        if (envelope) this.seen.set(envelope.commandId, { signature });
        return { view: this.view() };
      }
    }
    const candidates = new Set<string>(this.allocations);
    let metadataDirty = false;
    const ranges: Selection[] = [];
    try {
      for (const command of commands) {
        this.check();
        metadataDirty ||=
          !(
            ['core.cells.paste', 'core.cells.pasteTile'].includes(command.type) &&
            !changesStyles(command)
          ) &&
          ![
            'core.cells.set',
            'core.cells.annotation',
            'core.cells.copy',
            'core.cells.fillDown',
            'core.cells.fillSeries',
            'core.cells.clear',
            'core.cells.replace',
            'core.sheet.sort',
          ].includes(command.type);
        const p = command.payload as any;
        if (p.sheetId) {
          const sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId)!;
          if (p.target) ranges.push({ sheetId: p.sheetId, ...p.target });
          if (p.range) ranges.push({ sheetId: p.sheetId, ...p.range });
          else if (p.cells?.length) {
            let startRow = Infinity,
              endRow = 0,
              startColumn = Infinity,
              endColumn = 0;
            for (const cell of p.cells) {
              startRow = Math.min(startRow, cell.row);
              endRow = Math.max(endRow, cell.row + 1);
              startColumn = Math.min(startColumn, cell.column);
              endColumn = Math.max(endColumn, cell.column + 1);
            }
            ranges.push({
              sheetId: p.sheetId,
              startRow,
              endRow,
              startColumn,
              endColumn,
            });
          } else if (sh)
            ranges.push({
              sheetId: sh.id,
              startRow: 0,
              startColumn: 0,
              endRow: sh.rowOrder.length,
              endColumn: sh.columnOrder.length,
            });
        }
        if (
          [
            'core.axis.insert',
            'core.axis.delete',
            'core.sheet.rename',
            'core.sheet.remove',
          ].includes(command.type)
        )
          await this.structural(command, next, metadata);
        else if (command.type === 'core.cells.paste')
          await this.pasteCells(command, next, metadata);
        else if (command.type === 'core.cells.pasteTile')
          await this.pasteTile(command, next, metadata);
        else if (command.type === 'core.cells.copySpecial')
          await this.copySpecial(command, next, metadata);
        else if (command.type === 'core.cells.fillDown')
          await this.fillDown(command, next, metadata);
        else if (command.type === 'core.cells.fillSeries')
          await this.fillSeries(command, next, metadata);
        else if (command.type === 'core.cells.style')
          await this.applyStyle(command, next, metadata);
        else if (command.type === 'core.cells.replace')
          await this.replaceText(command, next, metadata);
        else if (command.type === 'core.sheet.sort') await this.sort(command, next, metadata);
        else await this.applyLocal(command, next, metadata);
      }
      if (
        commands.length &&
        commands.every((command) => command.type === 'core.cells.replace') &&
        !this.changed.size
      ) {
        if (envelope) this.seen.set(envelope.commandId, { signature });
        return { view: this.view() };
      }
      await this.filters(next, metadata);
      next.revision++;
      next.meta = !metadataDirty ? old.meta : await this.allocate(encodeMetadata(metadata));
      const changes: History['changes'] = [];
      for (const a of this.changed)
        if (old.blocks[a] !== next.blocks[a])
          changes.push([a, old.blocks[a] ?? null, next.blocks[a] ?? null]);
      const bytes =
        changes.reduce((n, [, a, b]) => n + (a?.bytes ?? 0) + (b?.bytes ?? 0), 0) +
        (next.meta === old.meta
          ? 0
          : encodedBytes(encodeMetadata(metadata)) + encodedBytes(encodeMetadata(this.metadata)));
      assert(
        bytes <= 32 * 1024 * 1024,
        'LIMIT_EXCEEDED',
        'Change exceeds the 32 MiB undo budget; import as a new workbook',
      );
      const history: History = { before: old.meta, after: next.meta, changes, bytes, ranges };
      next.past.push(await this.allocate(history));
      for (const id of old.future) {
        const h = this.histories.get(id)!;
        candidates.add(id);
        candidates.add(h.before);
        candidates.add(h.after);
        for (const [, a, b] of h.changes) {
          if (a) candidates.add(a.id);
          if (b) candidates.add(b.id);
        }
      }
      next.future = [];
      next.historyBytes = old.historyBytes + bytes;
      while (next.past.length > 100 || next.historyBytes > 32 * 1024 * 1024) {
        const expired = next.past.shift()!;
        const h = this.histories.get(expired)!;
        next.historyBytes -= h.bytes;
        candidates.add(expired);
        candidates.add(h.before);
        candidates.add(h.after);
        for (const [, a, b] of h.changes) {
          if (a) candidates.add(a.id);
          if (b) candidates.add(b.id);
        }
      }
      await this.publish(next, metadata, metadataDirty ? undefined : ranges);
      await this.cleanup([...candidates, ...this.allocations]);
      for (const [address, pointer] of transaction.changed) {
        if (pointer) old.blocks[address] = pointer;
        else delete old.blocks[address];
      }
      this.head.blocks = old.blocks;
      const commit: Commit = {
        commandId: envelope?.commandId ?? uid('cmd'),
        previousRevision: old.revision,
        revision: next.revision,
        source: label,
        commands,
        changedRanges: ranges,
      };
      if (envelope) this.seen.set(envelope.commandId, { signature, commit });
      return { view: this.view(), commit };
    } catch (error) {
      await this.cleanup(this.allocations);
      throw error;
    }
  }
  private async pasteTile(command: Command, next: Head, metadata: Metadata) {
    const tile = command.payload as PasteTile;
    const sh = metadata.snapshot.sheets.find((sheet) => sheet.id === tile.sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    let units = validatePasteTile(sh, tile);
    const charge = (extra: number) => {
      units += extra;
      assert(
        units <= 32 * 1024 * 1024,
        'LIMIT_EXCEEDED',
        'Expanded clipboard content is too large',
      );
    };
    if (changesStyles(command)) metadata.snapshot.styles = { ...metadata.snapshot.styles };
    const offset = createTileOffsetter();
    const rowGroups = groupAxis(sh.rowOrder, tile.range.startRow, tile.range.endRow, BLOCK_ROWS);
    const columnGroups = groupAxis(
      sh.columnOrder,
      tile.range.startColumn,
      tile.range.endColumn,
      BLOCK_COLUMNS,
    );
    for (const rows of rowGroups.values())
      for (const columns of columnGroups.values()) {
        this.check();
        const address = blockAddress(tile.sheetId, rows[0].id, columns[0].id);
        const cells = structuredClone((await this.block(address, next)).cells);
        for (const row of rows)
          for (const column of columns) {
            const position = key(row.id, column.id);
            const target =
              cells[position] ??
              (cells[position] = { rowId: row.id, columnId: column.id, input: { type: 'blank' } });
            applyPastedCell(
              metadata.snapshot,
              target,
              tileCell(tile, row.index, column.index, charge, offset),
            );
          }
        await this.storeBlock(address, cells, next, metadata);
      }
  }
  private async copySpecial(command: Command, next: Head, metadata: Metadata) {
    const { sheetId, range, target, mode } = command.payload as {
      sheetId: string;
      range: Rect;
      target: Rect;
      mode: 'values' | 'formulas' | 'formats';
    };
    const sheet = metadata.snapshot.sheets.find((item) => item.id === sheetId);
    assert(sheet, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    checkBounds(sheet, range);
    checkBounds(sheet, target);
    const rows = range.endRow - range.startRow;
    const columns = range.endColumn - range.startColumn;
    assert(
      rows === target.endRow - target.startRow && columns === target.endColumn - target.startColumn,
      'INVALID_RANGE',
      'Copy ranges must have equal dimensions',
    );
    assert(rows * columns <= LIMITS.rangeCells, 'LIMIT_EXCEEDED', 'Copy exceeds the cell limit');
    assert(
      !sheet.merges.some((merge) => intersects(merge, range) || intersects(merge, target)),
      'INVALID_RANGE',
      'Unmerge before copying',
    );
    const calculate =
      mode === 'values'
        ? new AsyncCalculation(
            (id, row, column) => this.stagedCell(id, row, column, next, metadata),
            (name) =>
              metadata.snapshot.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())?.id,
            this.calculator.cacheBudget,
            undefined,
            () => metadata.snapshot.dateSystem,
          ).readSession()
        : undefined;
    const page: StagedReadPage = {};
    const cells: PasteCell[] = [];
    let bytes = 0;
    let visited = 0;
    for (let row = range.startRow; row < range.endRow; row++) {
      for (let column = range.startColumn; column < range.endColumn; column++) {
        this.check();
        if (++visited % 1024 === 0) {
          await yieldToTask();
          this.check();
        }
        const cell = await this.stagedCell(sheetId, row, column, next, metadata, page);
        const style = cell?.styleId ? metadata.snapshot.styles[cell.styleId] : {};
        const item = selectiveCopyCell(
          {
            row,
            column,
            ...(mode === 'values'
              ? {
                  value:
                    cell?.input.type === 'formula'
                      ? await calculate!(sheetId, row, column)
                      : cell
                        ? scalarValue(cell.input)
                        : null,
                }
              : {}),
            ...(mode === 'formulas' ? { input: cell?.input } : {}),
            ...(mode === 'formats'
              ? { style, numberFormat: cell?.numberFormat ?? style?.numberFormat }
              : {}),
          },
          mode,
          target.startRow - range.startRow,
          target.startColumn - range.startColumn,
        );
        bytes += JSON.stringify(item).length * 2;
        assert(bytes <= 32 * 1024 * 1024, 'LIMIT_EXCEEDED', 'Copy working set is too large');
        cells.push(item);
      }
    }
    const paste = validateCommand({ type: 'core.cells.paste', payload: { sheetId, cells } });
    await this.pasteCells(paste, next, metadata);
  }
  private async pasteCells(command: Command, next: Head, metadata: Metadata) {
    const { sheetId, cells: incoming } = command.payload as { sheetId: string; cells: PasteCell[] };
    const sh = metadata.snapshot.sheets.find((sheet) => sheet.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    const pages = new Map<string, PasteCell[]>();
    let startRow = Infinity,
      endRow = 0,
      startColumn = Infinity,
      endColumn = 0;
    for (const item of incoming) {
      assert(
        item.row < sh.rowOrder.length && item.column < sh.columnOrder.length,
        'INVALID_RANGE',
        'Range exceeds sheet dimensions',
      );
      startRow = Math.min(startRow, item.row);
      endRow = Math.max(endRow, item.row + 1);
      startColumn = Math.min(startColumn, item.column);
      endColumn = Math.max(endColumn, item.column + 1);
      const address = blockAddress(sheetId, sh.rowOrder[item.row], sh.columnOrder[item.column]);
      let page = pages.get(address);
      if (!page) pages.set(address, (page = []));
      page.push(item);
    }
    const merges = sh.merges.filter((merge) =>
      intersects(merge, { startRow, endRow, startColumn, endColumn }),
    );
    let checks = 0;
    for (const item of incoming)
      for (const merge of merges) {
        assert(++checks <= 2_000_000, 'LIMIT_EXCEEDED', 'Paste merge-check budget exceeded');
        assert(
          !(
            item.row >= merge.startRow &&
            item.row < merge.endRow &&
            item.column >= merge.startColumn &&
            item.column < merge.endColumn
          ),
          'INVALID_RANGE',
          'Unmerge before pasting cell inputs or formats',
        );
      }
    if (changesStyles(command)) metadata.snapshot.styles = { ...metadata.snapshot.styles };
    for (const [address, page] of pages) {
      this.check();
      const cells = structuredClone((await this.block(address, next)).cells);
      for (const item of page) {
        const rowId = sh.rowOrder[item.row],
          columnId = sh.columnOrder[item.column];
        const position = key(rowId, columnId);
        const target =
          cells[position] ?? (cells[position] = { rowId, columnId, input: { type: 'blank' } });
        applyPastedCell(metadata.snapshot, target, item);
      }
      await this.storeBlock(address, cells, next, metadata);
    }
  }
  private async fillSeries(command: Command, next: Head, metadata: Metadata) {
    const { sheetId, range, start, step, axis, calendar } = command.payload as {
      sheetId: string;
      range: Rect;
      start: number;
      step: number;
      axis: 'row' | 'column';
      calendar?: 'month' | 'year';
    };
    const sh = metadata.snapshot.sheets.find((sheet) => sheet.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    checkBounds(sh, range);
    assert(
      (range.endRow - range.startRow) * (range.endColumn - range.startColumn) <= LIMITS.rangeCells,
      'LIMIT_EXCEEDED',
      'Series exceeds the cell limit',
    );
    assert(
      !sh.merges.some((merge) => intersects(merge, range)),
      'INVALID_RANGE',
      'Unmerge before filling a series',
    );
    const count =
      axis === 'row' ? range.endRow - range.startRow : range.endColumn - range.startColumn;
    const values = calendar
      ? calendarSeries(start, step, count, calendar, metadata.snapshot.dateSystem)
      : numericSeries(start, step, count);
    const rowGroups = groupAxis(sh.rowOrder, range.startRow, range.endRow, BLOCK_ROWS);
    const columnGroups = groupAxis(
      sh.columnOrder,
      range.startColumn,
      range.endColumn,
      BLOCK_COLUMNS,
    );
    for (const rows of rowGroups.values()) {
      for (const columns of columnGroups.values()) {
        this.check();
        const address = blockAddress(sheetId, rows[0].id, columns[0].id);
        const cells = structuredClone((await this.block(address, next)).cells);
        for (const row of rows)
          for (const column of columns) {
            const position = key(row.id, column.id);
            const target =
              cells[position] ??
              (cells[position] = {
                rowId: row.id,
                columnId: column.id,
                input: { type: 'blank' },
              });
            target.input = {
              type: 'number',
              value:
                values[
                  axis === 'row' ? row.index - range.startRow : column.index - range.startColumn
                ],
            };
            delete target.cached;
          }
        await this.storeBlock(address, cells, next, metadata);
      }
    }
  }
  private async applyStyle(command: Command, next: Head, metadata: Metadata) {
    const { sheetId, range, style } = command.payload as {
      sheetId: string;
      range: Rect;
      style: CellStyle;
    };
    const sh = metadata.snapshot.sheets.find((sheet) => sheet.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    checkBounds(sh, range);
    assert(
      (range.endRow - range.startRow) * (range.endColumn - range.startColumn) <= LIMITS.rangeCells,
      'LIMIT_EXCEEDED',
      'Style range exceeds the cell limit',
    );
    assert(
      sh.merges.every((merge) => !intersects(merge, range) || contains(range, merge)),
      'INVALID_RANGE',
      'Operation cuts through a merged cell',
    );
    metadata.snapshot.styles = { ...metadata.snapshot.styles };
    const styles = new Map<string, string>();
    const rowGroups = groupAxis(sh.rowOrder, range.startRow, range.endRow, BLOCK_ROWS);
    const columnGroups = groupAxis(
      sh.columnOrder,
      range.startColumn,
      range.endColumn,
      BLOCK_COLUMNS,
    );
    for (const rows of rowGroups.values()) {
      for (const columns of columnGroups.values()) {
        this.check();
        const address = blockAddress(sheetId, rows[0].id, columns[0].id);
        const cells = structuredClone((await this.block(address, next)).cells);
        for (const row of rows)
          for (const column of columns) {
            const position = key(row.id, column.id);
            const target =
              cells[position] ??
              (cells[position] = {
                rowId: row.id,
                columnId: column.id,
                input: { type: 'blank' },
              });
            const prior = target.styleId ?? '';
            let nextStyle = styles.get(prior);
            if (!nextStyle) {
              nextStyle = styleId(metadata.snapshot, {
                ...(target.styleId ? metadata.snapshot.styles[target.styleId] : {}),
                ...style,
              });
              styles.set(prior, nextStyle);
            }
            target.styleId = nextStyle;
            if (style.numberFormat !== undefined) delete target.numberFormat;
          }
        await this.storeBlock(address, cells, next, metadata);
      }
    }
  }
  private async fillDown(command: Command, next: Head, metadata: Metadata) {
    const { sheetId, range } = command.payload as { sheetId: string; range: Rect };
    const sh = metadata.snapshot.sheets.find((sheet) => sheet.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    checkBounds(sh, range);
    assert(
      !sh.merges.some((merge) => intersects(merge, range)),
      'INVALID_RANGE',
      'Unmerge before copying or filling',
    );
    const rowGroups = new Map<number, { index: number; id: number }[]>();
    const columnGroups = new Map<number, { index: number; id: number }[]>();
    for (let index = range.startRow + 1; index < range.endRow; index++) {
      const id = sh.rowOrder[index],
        group = Math.floor(id / BLOCK_ROWS);
      let values = rowGroups.get(group);
      if (!values) rowGroups.set(group, (values = []));
      values.push({ index, id });
    }
    for (let index = range.startColumn; index < range.endColumn; index++) {
      const id = sh.columnOrder[index],
        group = Math.floor(id / BLOCK_COLUMNS);
      let values = columnGroups.get(group);
      if (!values) columnGroups.set(group, (values = []));
      values.push({ index, id });
    }
    const source = new Map<number, CellRecord>();
    const sourceId = sh.rowOrder[range.startRow];
    for (const columns of columnGroups.values()) {
      this.check();
      const block = await this.block(blockAddress(sheetId, sourceId, columns[0].id), next);
      const cells = block.cells;
      for (const column of columns) {
        const value = cells[key(sourceId, column.id)];
        if (value) source.set(column.id, structuredClone(value));
      }
    }
    for (const rows of rowGroups.values()) {
      for (const columns of columnGroups.values()) {
        this.check();
        const address = blockAddress(sheetId, rows[0].id, columns[0].id);
        const cells = structuredClone((await this.block(address, next)).cells);
        for (const row of rows)
          for (const column of columns) {
            const position = key(row.id, column.id);
            const original = source.get(column.id);
            if (!original) {
              delete cells[position];
              continue;
            }
            const value = structuredClone(original);
            value.rowId = row.id;
            value.columnId = column.id;
            if (value.input.type === 'formula')
              value.input.expression = offsetCellFormula(
                value.input.expression,
                row.index - range.startRow,
                0,
              );
            delete value.cached;
            cells[position] = value;
          }
        await this.storeBlock(address, cells, next, metadata);
      }
    }
  }
  private async replaceText(command: Command, next: Head, metadata: Metadata) {
    const id = String(command.payload.sheetId);
    assert(
      metadata.snapshot.sheets.some((sheet) => sheet.id === id),
      'INVALID_ARGUMENT',
      'Unknown sheet',
    );
    const query = String(command.payload.query);
    if (query === command.payload.replacement) return;
    for (const a of Object.keys(next.blocks))
      if (a.startsWith(`${id}/`)) {
        this.check();
        const existing = (await this.block(a, next)).cells;
        if (
          !Object.values(existing).some(
            (cell) => cell.input.type === 'string' && cell.input.value.includes(query),
          )
        )
          continue;
        const cells = structuredClone(existing);
        const snapshot = {
          ...metadata.snapshot,
          sheets: metadata.snapshot.sheets.map((s) => (s.id === id ? { ...s, cells } : s)),
        };
        reduceCommand(snapshot, command);
        await this.storeBlock(a, cells, next, metadata);
      }
  }
  private async structural(command: Command, next: Head, metadata: Metadata) {
    const original = metadata.snapshot;
    const transformed = structuredClone(original);
    // Axes are copied only for a structural operation. Stable identity pages stay in place.
    reduceCommand(transformed, command);
    const axisPayload = command.payload as any;
    if (command.type === 'core.axis.insert' || command.type === 'core.axis.delete') {
      const before = original.sheets.find((s) => s.id === axisPayload.sheetId)!,
        after = transformed.sheets.find((s) => s.id === axisPayload.sheetId)!;
      const field = axisPayload.axis === 'row' ? 'rowOrder' : 'columnOrder',
        index = this.axis(before[field]).clone();
      index.splice(
        axisPayload.index,
        command.type === 'core.axis.delete' ? axisPayload.count : 0,
        command.type === 'core.axis.insert' ? axisPayload.ids : [],
      );
      assert(
        index.toArray().every((id, i) => id === after[field][i]),
        'CORRUPT_STORAGE',
        'Axis sequence differs',
      );
      this.axes.set(after[field], index);
    }
    const payload = command.payload as any,
      sourceSheet = original.sheets.find((s) => s.id === payload.sheetId)!;
    const sourceName = sourceSheet.name.toLowerCase();
    const names = new Map(original.sheets.map((sheet) => [sheet.id, sheet.name]));
    const rewrite = createCellFormulaTransformer(
      structuralReferences(command, sourceSheet.name),
      structuralRangeReferences(command, sourceSheet.name),
    );
    const axisEdit = command.type === 'core.axis.insert' || command.type === 'core.axis.delete';
    const deleted =
      command.type === 'core.axis.delete'
        ? new Set(
            (payload.axis === 'row' ? sourceSheet.rowOrder : sourceSheet.columnOrder).slice(
              payload.index,
              payload.index + payload.count,
            ),
          )
        : new Set<number>();
    const deletedBuckets = new Set(
      [...deleted].map((id) => Math.floor(id / (payload.axis === 'row' ? 64 : 32))),
    );
    const remainingBuckets = new Set<number>();
    if (command.type === 'core.axis.delete') {
      const sheet = transformed.sheets.find((item) => item.id === payload.sheetId)!;
      const order = payload.axis === 'row' ? sheet.rowOrder : sheet.columnOrder;
      const size = payload.axis === 'row' ? BLOCK_ROWS : BLOCK_COLUMNS;
      for (const id of order) {
        const bucket = Math.floor(id / size);
        if (deletedBuckets.has(bucket)) remainingBuckets.add(bucket);
        if (remainingBuckets.size === deletedBuckets.size) break;
      }
    }
    metadata.snapshot = transformed;
    for (const [address, pointer] of Object.entries(next.blocks)) {
      this.check();
      const [sheetId, br, bc] = address.split('/');
      const removedSheet = command.type === 'core.sheet.remove' && sheetId === sourceSheet.id;
      const deletion =
        sheetId === sourceSheet.id && deletedBuckets.has(payload.axis === 'row' ? +br : +bc);
      const formula = pointer.formulas?.some(
        (name) => name === '*' || name.toLowerCase() === sourceName,
      );
      if (removedSheet) {
        this.changed.add(address);
        delete next.blocks[address];
        continue;
      }
      if (deletion && !remainingBuckets.has(payload.axis === 'row' ? +br : +bc)) {
        this.changed.add(address);
        delete next.blocks[address];
        continue;
      }
      if (!deletion && !formula) continue;
      const cells = structuredClone((await this.block(address, next)).cells);
      if (deletion) {
        const field = payload.axis === 'row' ? 'rowId' : 'columnId';
        for (const [key, cell] of Object.entries(cells))
          if (deleted.has(cell[field])) delete cells[key];
      }
      rewrite(cells, axisEdit ? names.get(sheetId) : undefined);
      await this.storeBlock(address, cells, next, metadata);
    }
  }
  private async sort(command: Command, next: Head, metadata: Metadata) {
    const p = command.payload as any,
      sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId)!;
    checkBounds(sh, p.range);
    // Validate range/merge constraints using the existing semantic reducer.
    const keys = sortColumns(p, p.range);
    assert(
      !sh.merges.some((m) => intersects(m, p.range)),
      'INVALID_RANGE',
      'Unmerge before sorting',
    );
    const source = { ...next, blocks: { ...next.blocks } };
    for (const block of Object.values(source.blocks)) this.pins.add(block.id);
    let runs: Array<{ id: string; count: number; bytes: number; maxRowBytes: number }> = [];
    let scratchBytes = 0;
    let outputComplete = false;
    type SortRow = { row: number; cells: Array<CellRecord | null>; values: CellValue[] };
    const compare = (a: SortRow, b: SortRow) =>
      compareSortKeys(a.values, b.values, keys) || a.row - b.row;
    const calculationBytes = Math.min(4 * 1024 * 1024, this.cache.budget);
    const workingBudget = 32 * 1024 * 1024 - calculationBytes;
    const calculator = new AsyncCalculation(
      (sheet, row, column) => this.stagedCell(sheet, row, column, source, metadata),
      (name) =>
        metadata.snapshot.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())?.id,
      calculationBytes,
      undefined,
      () => metadata.snapshot.dateSystem,
    );
    try {
      let stageStart = performance.now();
      const totalRows = p.range.endRow - p.range.startRow;
      const runRows = Math.max(
        1,
        Math.min(256, Math.floor(8192 / (p.range.endColumn - p.range.startColumn))),
      );
      for (let row = p.range.startRow; row < p.range.endRow; row += runRows) {
        const batch: SortRow[] = [];
        let workingBytes = 0;
        for (let r = row; r < Math.min(row + runRows, p.range.endRow); r++) {
          this.check();
          const cells: Array<CellRecord | null> = [];
          for (let c = p.range.startColumn; c < p.range.endColumn; c++) {
            const cell =
              (
                await this.block(blockAddress(sh.id, sh.rowOrder[r], sh.columnOrder[c]), source)
              ).cell((sh.rowOrder[r] % 64) * 32 + (sh.columnOrder[c] % 32)) ?? null;
            workingBytes += encodedBytes(cell) * 4;
            assert(
              workingBytes <= workingBudget,
              'LIMIT_EXCEEDED',
              'Sort run exceeds the operation working-set budget',
            );
            cells.push(cell);
          }
          const values: CellValue[] = [];
          for (const key of keys) {
            const input = cells[key.column - p.range.startColumn]?.input;
            const value = sortKey(
              input?.type === 'formula'
                ? await calculator.value(sh.id, r, key.column)
                : input
                  ? scalarValue(input)
                  : null,
            );
            workingBytes += encodedBytes(value) * 4;
            assert(
              workingBytes <= workingBudget,
              'LIMIT_EXCEEDED',
              'Sort keys exceed the operation working-set budget',
            );
            values.push(value);
          }
          batch.push({ row: r, cells, values });
        }
        assert(
          encodedBytes(batch) * 4 <= workingBudget,
          'LIMIT_EXCEEDED',
          'Sort run exceeds the 32 MiB operation budget',
        );
        batch.sort(compare);
        const run = `${this.id}/sort/0/${runs.length}`;
        const puts: Array<[IDBValidKey, unknown]> = batch.map((v, i) => [
          `${run}/${String(i).padStart(8, '0')}`,
          v,
        ]);
        const runBytes = puts.reduce((total, record) => total + encodedBytes(record), 0);
        scratchBytes += runBytes;
        this.stagingBytes = scratchBytes;
        assert(
          this.usedBytes + scratchBytes <= next.budget,
          'STORAGE_BUDGET',
          'Sort staging exceeds the storage budget',
        );
        await this.db.write('scratch', puts);
        runs.push({
          id: run,
          count: batch.length,
          bytes: runBytes,
          maxRowBytes: Math.max(...batch.map((row) => encodedBytes(row) * 4 + 1024)),
        });
      }
      this.progress({
        phase: 'sort-runs',
        completed: totalRows,
        total: totalRows,
        elapsedMs: performance.now() - stageStart,
      });
      stageStart = performance.now();
      // Pairwise merge buffers are sized by the largest persisted row, not just row count.
      let pass = 1;
      while (runs.length > 1) {
        const merged: typeof runs = [];
        for (let i = 0; i < runs.length; i += 2) {
          if (!runs[i + 1]) {
            merged.push(runs[i]);
            continue;
          }
          const left = runs[i],
            right = runs[i + 1],
            id = `${this.id}/sort/${pass}/${i}`;
          const maxRowBytes = Math.max(left.maxRowBytes, right.maxRowBytes);
          assert(
            left.maxRowBytes + right.maxRowBytes <= workingBudget,
            'LIMIT_EXCEEDED',
            'Sort merge rows exceed the operation working-set budget',
          );
          // Reserve six buffer shares for both readers, pending output and transient
          // cursor/encoding copies. Oversized rows retain the two-row merge path.
          const buffered = maxRowBytes <= workingBudget / 6;
          const batchSize = buffered
            ? Math.min(64, Math.floor(workingBudget / 6 / maxRowBytes))
            : 1;
          const db = this.db;
          async function* rows(run: (typeof runs)[number]) {
            if (batchSize === 1) {
              for (let index = 0; index < run.count; index++) {
                const value = await db.get<SortRow>(
                  'scratch',
                  `${run.id}/${String(index).padStart(8, '0')}`,
                );
                assert(value, 'CORRUPT_STORAGE', 'Missing sort row');
                yield value;
              }
            } else {
              for await (const entries of db.scan<SortRow>(
                'scratch',
                prefixRange(run.id),
                batchSize,
              ))
                for (const [, value] of entries) yield value;
            }
          }
          const leftRows = rows(left),
            rightRows = rows(right);
          let x = await leftRows.next(),
            y = await rightRows.next();
          let output = 0,
            bytes = 0;
          let pending: Array<[IDBValidKey, unknown]> = [];
          const flushMerge = async () => {
            if (pending.length) await this.db.write('scratch', pending);
            pending = [];
          };
          try {
            while (!x.done || !y.done) {
              this.check();
              const takeLeft = !x.done && (y.done || compare(x.value, y.value) <= 0);
              const value = takeLeft ? x.value! : y.value!;
              const record: [IDBValidKey, unknown] = [
                `${id}/${String(output++).padStart(8, '0')}`,
                value,
              ];
              const size = encodedBytes(record);
              scratchBytes += size;
              bytes += size;
              this.stagingBytes = scratchBytes;
              assert(
                this.usedBytes + scratchBytes <= next.budget,
                'STORAGE_BUDGET',
                'Sort merge exceeds the storage budget',
              );
              pending.push(record);
              if (pending.length >= batchSize) await flushMerge();
              if (takeLeft) x = await leftRows.next();
              else y = await rightRows.next();
            }
            await flushMerge();
            assert(output === left.count + right.count, 'CORRUPT_STORAGE', 'Incomplete sort run');
          } finally {
            await leftRows.return(undefined);
            await rightRows.return(undefined);
          }
          await this.db.write('scratch', [], [prefixRange(left.id), prefixRange(right.id)]);
          scratchBytes -= left.bytes + right.bytes;
          this.stagingBytes = scratchBytes;
          merged.push({ id, count: output, bytes, maxRowBytes });
        }
        runs = merged;
        pass++;
      }
      this.progress({
        phase: 'sort-merge',
        completed: totalRows,
        total: totalRows,
        elapsedMs: performance.now() - stageStart,
      });
      stageStart = performance.now();
      let position = 0;
      const grouped = new Map<string, Record<string, CellRecord>>();
      let groupedBytes = 0,
        rowBand = -1;
      const flush = async () => {
        for (const [address, cells] of grouped) {
          this.check();
          await this.storeBlock(address, cells, next, metadata);
        }
        grouped.clear();
        groupedBytes = 0;
      };
      for (let row = p.range.startRow; row < p.range.endRow; row++) {
        this.check();
        const entry = (await this.db.get<SortRow>(
          'scratch',
          `${runs[0].id}/${String(position++).padStart(8, '0')}`,
        ))!;
        const entryBytes = encodedBytes(entry) * 4;
        assert(
          entryBytes <= workingBudget,
          'LIMIT_EXCEEDED',
          'Sort output row exceeds the operation budget',
        );
        const band = Math.floor(sh.rowOrder[row] / 64);
        if (rowBand !== band || groupedBytes + entryBytes > workingBudget) await flush();
        rowBand = band;
        const pageFor = async (address: string, extra = 0) => {
          if (grouped.has(address) && groupedBytes + entryBytes + extra > workingBudget)
            await flush();
          if (!grouped.has(address)) {
            const cells = (await this.block(address, next)).cells;
            const bytes = encodedBytes(cells) * 4;
            if (groupedBytes + bytes + entryBytes + extra > workingBudget) await flush();
            assert(
              bytes + entryBytes + extra <= workingBudget,
              'LIMIT_EXCEEDED',
              'Sort output page exceeds the operation budget',
            );
            grouped.set(address, structuredClone(cells));
            groupedBytes += bytes;
          }
          return grouped.get(address)!;
        };
        for (let c = p.range.startColumn; c < p.range.endColumn; c++) {
          const a = blockAddress(sh.id, sh.rowOrder[row], sh.columnOrder[c]);
          let cells = await pageFor(a);
          const k = key(sh.rowOrder[row], sh.columnOrder[c]);
          const cell = entry.cells[c - p.range.startColumn];
          if (!cell) {
            if (cells[k]) groupedBytes -= encodedBytes(cells[k]) * 4;
            delete cells[k];
          } else {
            const copied = structuredClone(cell);
            copied.rowId = sh.rowOrder[row];
            copied.columnId = sh.columnOrder[c];
            delete copied.cached;
            if (copied.input.type === 'formula')
              copied.input.expression = offsetCellFormula(
                copied.input.expression,
                row - entry.row,
                0,
              );
            const growth =
              (encodedBytes(copied) -
                (cells[k] ? encodedBytes(cells[k]) : 0) +
                (cells[k] ? 0 : k.length + 4)) *
              4;
            cells = await pageFor(a, Math.max(0, growth));
            cells[k] = copied;
            groupedBytes += growth;
          }
        }
      }
      await flush();
      this.progress({
        phase: 'sort-output',
        completed: totalRows,
        total: totalRows,
        elapsedMs: performance.now() - stageStart,
      });
      outputComplete = true;
    } finally {
      calculator.clear();
      try {
        await this.db.write('scratch', [], [prefixRange(`${this.id}/sort`)]);
        this.stagingBytes = 0;
      } catch {
        this.poisoned = true;
        if (outputComplete)
          throw new OpenSheetError('WORKER_FAILED', 'Reopen the workbook to recover sort staging');
      }
    }
  }
  async history(redo: boolean) {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.cancelled = false;
    this.changed.clear();
    const next = {
        ...this.head,
        blocks: { ...this.head.blocks },
        past: [...this.head.past],
        future: [...this.head.future],
      },
      from = redo ? next.future : next.past,
      to = redo ? next.past : next.future;
    const id = from.pop();
    if (!id) return { view: this.view(), commit: undefined };
    const h = this.histories.get(id) ?? (await this.readHistory(id));
    assert(h, 'CORRUPT_STORAGE', 'Missing history');
    for (const [a, before, after] of h.changes) {
      this.changed.add(a);
      const p = redo ? after : before;
      if (p) next.blocks[a] = p;
      else delete next.blocks[a];
    }
    next.meta = redo ? h.after : h.before;
    next.revision++;
    to.push(id);
    next.historyBytes += redo ? h.bytes : -h.bytes;
    const metadata = {
      ...(next.meta === this.head.meta ? this.metadata : await this.readMetadata(next.meta)),
    };
    // Historical pages and metadata must agree before the new head is published.
    this.checkDirectorySheets(next, metadata);
    await this.filters(next, metadata);
    await this.publish(next, metadata, next.meta === this.head.meta ? h.ranges : undefined);
    return {
      view: this.view(),
      commit: {
        commandId: uid('cmd'),
        previousRevision: next.revision - 1,
        revision: next.revision,
        source: redo ? 'redo' : 'undo',
        commands: [],
        changedRanges: h.ranges,
      } as Commit,
    };
  }
  async search(
    sheetId: string,
    query: string,
    after = -1,
    cooperative = false,
    direction: 'next' | 'previous' = 'next',
  ) {
    this.check();
    validateSearchArguments(sheetId, query, after, direction);
    assert(
      typeof cooperative === 'boolean',
      'INVALID_ARGUMENT',
      'Invalid search cancellation mode',
    );
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId);
    assert(sh, 'INVALID_ARGUMENT', 'Unknown search sheet');
    const q = query.toLowerCase();
    const reverse = direction === 'previous';
    const precedes = (a: number, b: number) => (reverse ? a > b : a < b);
    const width = sh.columnOrder.length;
    const toResult = (position: number | undefined, count: number) => ({
      match:
        position === undefined
          ? undefined
          : { row: Math.floor(position / width), column: position % width },
      count,
    });
    const head = this.head;
    const metadata = this.metadata;
    const cached = this.searchResults.get(head);
    if (cached?.sheetId === sheetId && cached.query === q) {
      const { positions } = cached;
      let low = 0,
        high = positions.length;
      while (low < high) {
        const mid = (low + high) >>> 1;
        if (reverse ? positions[mid] < after : positions[mid] <= after) low = mid + 1;
        else high = mid;
      }
      return toResult(
        reverse ? (positions[low - 1] ?? positions.at(-1)) : (positions[low] ?? positions[0]),
        positions.length,
      );
    }
    let positions: number[] | undefined = q.length <= 1024 ? [] : undefined;
    // A zero-cache search would otherwise reread the same page for every
    // reference and persisted AST in a formula chain. Keep at most one
    // bounded page for this operation; it never enters the persistent cache.
    const page: StagedReadPage = {};
    const calculator = this.calculator.cacheBudget
      ? this.calculator
      : new AsyncCalculation(
          (id, row, column) => this.stagedCell(id, row, column, head, metadata, page),
          (name) =>
            metadata.snapshot.sheets.find((item) => item.name.toLowerCase() === name.toLowerCase())
              ?.id,
          0,
          async (id, row, column) => {
            const sheet = metadata.snapshot.sheets.find((item) => item.id === id)!;
            const rowId = sheet.rowOrder[row],
              columnId = sheet.columnOrder[column];
            return (await this.stagedPage(blockAddress(id, rowId, columnId), head, page))
              .dependencies[key(rowId, columnId)] as AST | null | undefined;
          },
          () => metadata.snapshot.dateSystem,
        );
    const calculate = this.calculator.cacheBudget
      ? (id: string, row: number, column: number) => calculator.value(id, row, column)
      : calculator.readSession(16 * 1024);
    let wrap: number | undefined,
      next: number | undefined,
      count = 0;
    const rows = this.axis(sh.rowOrder),
      cols = this.axis(sh.columnOrder);
    let visited = 0,
      textUnits = 0;
    for (const [a] of Object.entries(this.head.blocks))
      if (a.startsWith(`${sheetId}/`))
        for (const cell of (await this.block(a)).records()) {
          if (cooperative && ++visited % 16384 === 0) {
            await new Promise<void>((resolve) => setTimeout(resolve, 0));
            textUnits = 0;
          }
          this.check();
          const row = rows.get(cell.rowId)!,
            column = cols.get(cell.columnId)!;
          const value =
            cell.input.type === 'formula'
              ? await calculate(sheetId, row, column)
              : scalarValue(cell.input);
          assert(
            !(value && typeof value === 'object' && value.error === '#LIMIT!'),
            'LIMIT_EXCEEDED',
            'Search calculation exceeded its work budget',
          );
          const display =
            value &&
            typeof value === 'object' &&
            cell.cached &&
            this.head.revision === this.metadata.importRevision
              ? formatValue(
                  scalarValue(cell.cached),
                  cell.numberFormat,
                  this.metadata.snapshot.dateSystem,
                ) + ' †'
              : formatValue(
                  value,
                  cell.numberFormat ??
                    this.metadata.snapshot.styles[cell.styleId ?? '']?.numberFormat,
                  this.metadata.snapshot.dateSystem,
                );
          // Long strings can monopolize a Worker even when fewer than 16,384
          // records are scanned. Yield before normalizing the next large batch.
          if (cooperative) {
            textUnits += display.length;
            if (textUnits >= SEARCH_YIELD_TEXT_UNITS) {
              textUnits = 0;
              await new Promise<void>((resolve) => setTimeout(resolve, 0));
              this.check();
            }
          }
          if (!display.toLowerCase().includes(q)) continue;
          count++;
          const position = row * width + column;
          if (positions) {
            if (positions.length === 16384) positions = undefined;
            else positions.push(position);
          }
          if (wrap === undefined || precedes(position, wrap)) wrap = position;
          if (precedes(after, position) && (next === undefined || precedes(position, next)))
            next = position;
        }
    this.check();
    // Row-major positions fit uint32 under the workbook dimension limits.
    if (positions && this.head === head)
      this.searchResults.set(head, {
        sheetId,
        query: q,
        positions: Uint32Array.from(positions).sort(),
      });
    return toResult(next ?? wrap, count);
  }
  async clear() {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    await this.db.deleteWorkbook(this.id, prefixRange(this.id));
    this.cache.clear();
  }
  close() {
    this.writer = false;
    this.release?.();
    this.release = undefined;
    this.db.close();
    this.cache.clear();
    this.seen.clear();
  }
}
// Stateless fixed-seed generation allows exact validation without retaining the fixture.
export function fixtureValue(n: number): number | string | boolean {
  let x = (Math.imul(n + 1, 1664525) + 1013904223) >>> 0;
  if (n % 10 < 6) return n === 0 ? 0 : x / 1000;
  if (n % 10 === 9) return ((n / 10) | 0) % 2 === 1;
  let text = '';
  for (let i = 0; i < 32; i++) {
    x = (Math.imul(x, 1664525) + 1013904223) >>> 0;
    text += String.fromCharCode(33 + (x % 94));
  }
  return text;
}
