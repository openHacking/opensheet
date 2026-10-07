import { AxisIndex } from './axis-index.js';
import { overlay } from './copy-on-write.js';
import { intersects } from './address.js';
import { parseFormula, offsetFormula, type AST } from '@opensheetjs/formula';
import { decodeBlock, encodeBlock, ColumnPage, type EncodedBlock } from './block-codec.js';
import {
  decodeFileMetadata,
  encodeWorkbookFile,
  decodeFileCells,
  encodeFileCells,
  type WorkbookFile,
} from './file-codec.js';
import { crc32, gzipPage, unpackPage, readBinary } from './binary-file.js';
import { parseWorkbookJSON, type JSONSource } from './json-stream.js';
import { AsyncCalculation } from './async-calculation.js';
import { reduceCommand, validateCommand } from './commands.js';
import { formatValue } from './format.js';
import { checkBounds, validateSnapshot } from './model.js';
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
  key,
  uid,
  scalarValue,
  type CellRecord,
  type CellValue,
  type Command,
  type CommandEnvelope,
  type Commit,
  type Rect,
  type Selection,
  type WorkbookSnapshot,
} from './types.js';
export type Block = { cells: Record<string, CellRecord>; dependencies: Record<string, unknown> };
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
export type Progress = { phase: string; completed: number; total?: number };
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
  private compressionQueue = new Set<string>();
  private directoryEntries = new Map<string, Record<string, Pointer>>();
  private changed = new Set<string>();
  private references = new Map<string, number>();
  private histories = new Map<string, History>();
  private used: Record<string, Rect> = {};
  private sizes = new Map<string, number>();
  private cache: ByteCache<ColumnPage>;
  private cancelled = false;
  private writer = false;
  private release?: () => void;
  private calculator: AsyncCalculation;
  private seen: ByteCache<{ signature: string; commit: Commit }>;
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
    );
  }
  get isWriter() {
    return this.writer;
  }
  async lock() {
    if (typeof navigator !== 'undefined' && navigator.locks) {
      await new Promise<void>((resolve, reject) => {
        void navigator.locks
          .request(
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
    else throw new Error('Web Locks is required for safe multi-tab editing');
  }
  private check() {
    assert(!this.cancelled, 'ABORTED', 'Operation cancelled');
  }
  begin() {
    this.cancelled = false;
  }
  cancel() {
    this.cancelled = true;
  }
  async open(input?: WorkbookSnapshot | WorkbookFile, budget = 256 * 1024 * 1024) {
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
    const storedHead = await this.db.get<Head>('heads', this.id);
    const existing = storedHead ? await this.restoreHead(storedHead) : undefined;
    if (existing) {
      this.head = existing;
      this.metadata = await this.readMetadata(this.head.meta);
      assert(this.metadata, 'CORRUPT_STORAGE', 'Workbook metadata is missing');
      if (this.writer) await this.recover();
      else {
        this.usedBytes = await this.liveBytes();
        await this.initializeReferences();
      }
    } else {
      assert(snapshot && this.writer, 'INVALID_ARGUMENT', 'Workbook does not exist or is locked');
      // A worker can die during initial creation before any head exists.
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
      const h = await this.db.get<History>('records', id);
      if (!h) continue;
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
    for (const id of await this.liveIds()) {
      const bytes = await this.db.get<number>('sizes', id);
      if (bytes !== undefined) total += bytes + recordBytes(id, bytes);
    }
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
  private async initializeReferences() {
    this.references.clear();
    this.histories.clear();
    this.retain(this.head.meta, 1);
    for (const id of Object.values(this.head.directory ?? {})) this.retain(id, 1);
    for (const p of Object.values(this.head.blocks)) this.retain(p.id, 1);
    for (const id of [...this.head.past, ...this.head.future]) {
      const h = await this.db.get<History>('records', id);
      assert(h, 'CORRUPT_STORAGE', 'Missing history');
      this.histories.set(id, h);
      this.retainHistory(id, h, 1);
    }
  }
  private async recover() {
    await this.initializeReferences();
    for (const id of this.references.keys()) this.compressionQueue.add(id);
    const live = await this.liveIds();
    this.usedBytes = recordBytes(this.id, this.storedHead(this.head));
    this.sizes.clear();
    for await (const batch of this.db.scan<number>('sizes', prefixRange(this.id))) {
      const dead: string[] = [];
      for (const [k, v] of batch) {
        const id = String(k);
        if (!live.has(id)) dead.push(id);
        else {
          const bytes = v + recordBytes(id, v);
          this.usedBytes += bytes;
          this.sizes.set(id, bytes);
        }
      }
      if (dead.length) await this.db.write('records', [], dead);
    }
    await this.db.write('scratch', [], [prefixRange(this.id)]);
  }
  async reload() {
    this.head = await this.restoreHead((await this.db.get<Head>('heads', this.id))!);
    this.metadata = await this.readMetadata(this.head.meta);
    this.exportAddresses = undefined;
    this.cache.clear();
    this.calculator.clear();
    this.usedBytes = await this.liveBytes();
    await this.initializeReferences();
    this.used = {};
    return this.view();
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
  private async readMetadata(id: string): Promise<Metadata> {
    const stored = await this.db.get<StoredMetadata>('records', id);
    assert(stored, 'CORRUPT_STORAGE', 'Workbook metadata is missing');
    const snapshot = decodeFileMetadata(stored.snapshot);
    assert(
      Array.isArray(stored.styleIds) &&
        stored.styleIds.length === stored.snapshot.styles.length &&
        new Set(stored.styleIds).size === stored.styleIds.length &&
        stored.styleIds.every(
          (id) => /^[-\w]+$/.test(id) && !['__proto__', 'prototype', 'constructor'].includes(id),
        ),
      'CORRUPT_STORAGE',
      'Invalid stored style identities',
    );
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
    const dead = [...new Set(candidates)].filter((id) => !this.references.has(id));
    if (dead.length) await this.db.write('records', [], dead);
    for (const id of dead) {
      this.usedBytes -= this.sizes.get(id) ?? 0;
      this.sizes.delete(id);
      this.cache.delete(id);
    }
  }
  private async cleanup(candidates: Iterable<string>) {
    // A durable head is already committed. Failed reclamation is retried on reopen.
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
  private async restoreHead(head: Head) {
    assert(
      head.directory && !Object.hasOwn(head, 'blocks'),
      'CORRUPT_STORAGE',
      'Unsupported workbook database format',
    );
    const blocks: Record<string, Pointer> = {};
    this.directoryEntries.clear();
    for (const [bucket, id] of Object.entries(head.directory)) {
      const entries = await this.db.get<Record<string, Pointer>>('records', id);
      assert(entries, 'CORRUPT_STORAGE', 'Missing block directory');
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
    const bytes =
      this.usedBytes -
      recordBytes(this.id, this.storedHead(this.head)) +
      recordBytes(this.id, this.storedHead(next));
    assert(bytes <= next.budget, 'STORAGE_BUDGET', 'Application storage budget exceeded');
    await this.db.write('heads', [[this.id, this.storedHead(next)]]);
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
        const h = this.histories.get(id) ?? (await this.db.get<History>('records', id));
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
  }
  view(): EngineView {
    const used = this.used;
    for (const s of this.metadata.snapshot.sheets) {
      if (used[s.id]) continue;
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
  private async block(address: string, head = this.head): Promise<ColumnPage> {
    const p = head.blocks[address];
    if (!p) return decodeBlock(encodeBlock({}, [], [], {}));
    const cached = this.cache.get(p.id);
    if (cached) return cached;
    const encoded = await this.db.get<EncodedBlock>('records', p.id);
    assert(encoded, 'CORRUPT_STORAGE', 'Data block is missing');
    const block = decodeBlock(await this.inflate(encoded));
    this.cache.set(p.id, block, block.bytes);
    return block;
  }
  async cell(sheet: string, row: number, column: number): Promise<CellRecord | undefined> {
    this.check();
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheet);
    if (
      !sh ||
      row < 0 ||
      column < 0 ||
      row >= sh.rowOrder.length ||
      column >= sh.columnOrder.length
    )
      return { rowId: 0, columnId: 0, input: { type: 'error', code: '#REF!' } };
    return (await this.block(blockAddress(sheet, sh.rowOrder[row], sh.columnOrder[column]))).cell(
      (sh.rowOrder[row] % 64) * 32 + (sh.columnOrder[column] % 32),
    );
  }
  async read(sheetId: string, range: Rect, values = true, render = true) {
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
    for (let row = range.startRow; row < range.endRow; row++)
      for (let column = range.startColumn; column < range.endColumn; column++) {
        this.check();
        if ((row * sh.columnOrder.length + column) % 1024 === 0)
          await new Promise<void>((resolve) => setTimeout(resolve, 0));
        const k = key(sh.rowOrder[row], sh.columnOrder[column]);
        const cell = await this.cell(sheetId, row, column);
        if (cell) {
          cells[k] = cell;
          if (this.head.revision !== this.metadata.importRevision && cell.cached) {
            cells[k] = { ...cell };
            delete cells[k].cached;
          }
        }
        if (values) {
          const v = await this.calculator.value(sheetId, row, column);
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
    const dependencies: Record<string, unknown> = {};
    let count = 0,
      endRow = 1,
      endColumn = 1;
    for (const [k, c] of Object.entries(cells)) {
      if (c.input.type === 'blank' && !c.styleId && !c.numberFormat && !c.link && !c.note) {
        delete cells[k];
        continue;
      }
      if (c.input.type !== 'blank') count++;
      endRow = Math.max(endRow, rows.get(c.rowId)! + 1);
      endColumn = Math.max(endColumn, cols.get(c.columnId)! + 1);
      if (c.input.type === 'formula') {
        try {
          dependencies[k] = parseFormula(c.input.expression);
        } catch {
          dependencies[k] = null;
        }
      }
    }
    if (!Object.keys(cells).length) {
      delete head.blocks[address];
      await discardPrevious();
      return;
    }
    const block: Block = { cells, dependencies };
    const [, br, bc] = address.split('/');
    const encoded = encodeBlock(
      cells,
      Object.values(cells).map((c) => c.rowId),
      Object.values(cells).map((c) => c.columnId),
      dependencies,
    );
    const id = await this.allocate(encoded, head.budget);
    this.compressionQueue.add(id);
    let rowMask = 0n,
      columnMask = 0;
    const formulas = new Set<string>();
    for (const cell of Object.values(cells)) {
      rowMask |= 1n << BigInt(cell.rowId % 64);
      columnMask |= 1 << cell.columnId % 32;
      if (cell.input.type === 'formula') {
        const ast = dependencies[key(cell.rowId, cell.columnId)] as AST | null;
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
        if (ast) visit(ast);
        else formulas.add('*');
      }
    }
    head.blocks[address] = {
      id,
      bytes: this.sizes.get(id)!,
      count,
      endRow,
      endColumn,
      rowMask: rowMask.toString(16),
      columnMask: columnMask >>> 0,
      ...(formulas.size ? { formulas: [...formulas] } : {}),
      ...(Object.values(cells).some((c) => c.styleId !== undefined)
        ? {
            styles: [
              ...new Set(
                Object.values(cells).flatMap((c) => (c.styleId === undefined ? [] : [c.styleId])),
              ),
            ],
          }
        : {}),
    };
    const page = decodeBlock(encoded);
    this.cache.set(id, page, page.bytes);
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
      await this.collect(this.allocations);
      throw error;
    }
  }
  async importJSON(input: WorkbookFile | JSONSource, budget = this.head.budget) {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.allocations = [];
    this.changed.clear();
    this.allocated.clear();
    this.pins.clear();
    const old = await this.liveIds();
    for (const address of Object.keys(this.head.blocks)) this.changed.add(address);
    let file: WorkbookFile | undefined,
      staged = 0,
      published = false;
    const scratch = `${this.id}/import`;
    const stage = async (sheet: number, block: unknown) => {
      this.check();
      const record: [IDBValidKey, unknown] = [
        `${scratch}/${String(staged++).padStart(10, '0')}`,
        { sheet, block },
      ];
      this.stagingBytes += recordBytes(record[0], record[1]);
      assert(
        this.usedBytes + this.stagingBytes <= budget,
        'STORAGE_BUDGET',
        'Import staging exceeds storage budget',
      );
      await this.db.write('scratch', [record]);
      this.progress({ phase: 'import-parse', completed: staged });
    };
    try {
      if (input && typeof input === 'object' && 'binary' in input) {
        for await (const item of readBinary(
          (input as unknown as { binary: Blob | Uint8Array | ReadableStream<Uint8Array> }).binary,
        )) {
          if ('metadata' in item) file = item.metadata;
          else {
            const page = decodeBlock(item.bytes);
            assert(
              page.row === item.row && page.column === item.column,
              'INVALID_ARGUMENT',
              'Binary page coordinates differ',
            );
            const styles = new Map(file!.styles.map((_, i) => [`s${i}`, i]));
            await stage(item.sheet, {
              row: item.row,
              column: item.column,
              data: encodeFileCells(Object.values(page.cells), styles),
            });
          }
        }
      } else if (
        typeof input === 'string' ||
        input instanceof Blob ||
        input instanceof ReadableStream
      ) {
        for await (const item of parseWorkbookJSON(input)) {
          if ('block' in item) await stage(item.sheet, item.block);
          else file = item.metadata;
        }
      } else {
        assert(input && Array.isArray(input.sheets), 'INVALID_ARGUMENT', 'Invalid workbook file');
        file = { ...input, sheets: input.sheets.map((sh) => ({ ...sh, blocks: [] })) };
        for (let i = 0; i < input.sheets.length; i++)
          for (const block of input.sheets[i].blocks) await stage(i, block);
      }
      this.check();
      assert(file, 'INVALID_ARGUMENT', 'Missing workbook metadata');
      const snapshot = decodeFileMetadata(file);
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
        rows = snapshot.sheets.map((sh) => new Set(sh.rowOrder)),
        columns = snapshot.sheets.map((sh) => new Set(sh.columnOrder));
      let completed = 0;
      const seen = new Set<string>();
      for await (const batch of this.db.scan<{
        sheet: number;
        block: import('./file-codec.js').FileBlock;
      }>('scratch', prefixRange(scratch), 1)) {
        const [id, item] = batch[0];
        this.check();
        assert(
          Number.isInteger(item.sheet) && !!snapshot.sheets[item.sheet],
          'INVALID_ARGUMENT',
          'Unknown imported sheet',
        );
        const sh = snapshot.sheets[item.sheet],
          address = `${sh.id}/${item.block.row}/${item.block.column}`;
        assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate imported block');
        seen.add(address);
        const cells = decodeFileCells(item.block, styles, rows[item.sheet], columns[item.sheet]);
        await this.storeBlock(address, cells, next, metadata);
        await this.db.write('scratch', [], [id]);
        this.stagingBytes -= recordBytes(id, item);
        this.progress({ phase: 'import-store', completed: ++completed, total: staged });
      }
      await this.filters(next, metadata);
      next.meta = await this.allocate(encodeMetadata(metadata), budget);
      await this.publish(next, metadata);
      published = true;
      await this.cleanup([...old, ...this.allocations]);
      return this.view();
    } catch (error) {
      await this.collect(this.allocations);
      throw error;
    } finally {
      try {
        await this.db.write('scratch', [], [prefixRange(scratch)]);
      } catch (error) {
        if (!published) throw error;
      }
      this.stagingBytes = 0;
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
      await this.collect(this.allocations);
      throw error;
    }
  }
  private async filters(head: Head, metadata: Metadata) {
    metadata.filtered = {};
    for (const sh of metadata.snapshot.sheets)
      if (sh.filter) {
        const hidden: number[] = [],
          q = sh.filter.query.toLowerCase();
        // Temporarily evaluate staged data; the caller restores committed state on failure.
        const originalHead = this.head,
          originalMetadata = this.metadata;
        this.head = head;
        this.metadata = metadata;
        this.calculator.clear();
        try {
          for (let row = 1; row < sh.rowOrder.length; row++) {
            this.check();
            const c = await this.cell(sh.id, row, sh.filter.column);
            const value = await this.calculator.value(sh.id, row, sh.filter.column);
            if (
              !formatValue(
                value,
                c?.numberFormat ?? metadata.snapshot.styles[c?.styleId ?? '']?.numberFormat,
                metadata.snapshot.dateSystem,
              )
                .toLowerCase()
                .includes(q)
            )
              hidden.push(row);
          }
        } finally {
          this.head = originalHead;
          this.metadata = originalMetadata;
          this.calculator.clear();
        }
        metadata.filtered[sh.id] = hidden;
      }
  }
  private addresses(sheetId: string, range: Rect) {
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId)!;
    const rows = new Set(
      sh.rowOrder.slice(range.startRow, range.endRow).map((id) => Math.floor(id / BLOCK_ROWS)),
    );
    const columns = new Set(
      sh.columnOrder
        .slice(range.startColumn, range.endColumn)
        .map((id) => Math.floor(id / BLOCK_COLUMNS)),
    );
    const out: string[] = [];
    for (const r of rows) for (const c of columns) out.push(`${sheetId}/${r}/${c}`);
    return out;
  }
  private async applyLocal(command: Command, head: Head, metadata: Metadata) {
    const p = command.payload as any,
      sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId);
    const traversalCells =
      (p.cells?.length ?? 0) +
      [p.range, p.target].reduce(
        (total, range) =>
          total +
          (range ? (range.endRow - range.startRow) * (range.endColumn - range.startColumn) : 0),
        0,
      );
    assert(
      traversalCells * 512 <= 32 * 1024 * 1024,
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
          for (const a of this.addresses(sh.id, range)) addresses.add(a);
        }
    }
    let bytes = 0;
    const snapshot: WorkbookSnapshot = {
      ...metadata.snapshot,
      sheetOrder: [...metadata.snapshot.sheetOrder],
      styles:
        command.type === 'core.cells.style'
          ? { ...metadata.snapshot.styles }
          : metadata.snapshot.styles,
      sheets: metadata.snapshot.sheets.map((s) => ({
        ...s,
        cells: {},
        rows: command.type === 'core.axis.meta' ? { ...s.rows } : s.rows,
        columns: command.type === 'core.axis.meta' ? { ...s.columns } : s.columns,
        merges: command.type === 'core.cells.merge' ? s.merges.map((m) => ({ ...m })) : s.merges,
        freeze: s.freeze,
        filter: s.filter,
      })),
    };
    const target = snapshot.sheets.find((s) => s.id === p.sheetId);
    for (const address of addresses) {
      const block = await this.block(address, head);
      bytes += encodedBytes(block) * 4;
      assert(
        bytes <= 32 * 1024 * 1024,
        'LIMIT_EXCEEDED',
        'Command working set exceeds 32 MiB; use a streamed operation',
      );
      Object.assign(target!.cells, structuredClone(block.cells));
    }
    reduceCommand(snapshot, command);
    for (const address of addresses) {
      const [, br, bc] = address.split('/');
      const cells: Record<string, CellRecord> = {};
      for (const cell of Object.values(target!.cells))
        if (Math.floor(cell.rowId / 64) === +br && Math.floor(cell.columnId / 32) === +bc)
          cells[key(cell.rowId, cell.columnId)] = cell;
      await this.storeBlock(address, cells, head, metadata);
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
    const candidates = new Set<string>(this.allocations);
    let metadataDirty = false;
    const ranges: Selection[] = [];
    try {
      for (const command of commands) {
        this.check();
        metadataDirty ||= ![
          'core.cells.set',
          'core.cells.copy',
          'core.cells.fillDown',
          'core.cells.clear',
          'core.cells.replace',
          'core.sheet.sort',
        ].includes(command.type);
        const p = command.payload as any;
        if (p.sheetId) {
          const sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId)!;
          if (p.target) ranges.push({ sheetId: p.sheetId, ...p.target });
          if (p.range) ranges.push({ sheetId: p.sheetId, ...p.range });
          else if (p.cells?.length)
            ranges.push({
              sheetId: p.sheetId,
              startRow: Math.min(...p.cells.map((c: any) => c.row)),
              endRow: Math.max(...p.cells.map((c: any) => c.row)) + 1,
              startColumn: Math.min(...p.cells.map((c: any) => c.column)),
              endColumn: Math.max(...p.cells.map((c: any) => c.column)) + 1,
            });
          else if (sh)
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
        else if (command.type === 'core.cells.replace')
          await this.replaceText(command, next, metadata);
        else if (command.type === 'core.sheet.sort') await this.sort(command, next, metadata);
        else await this.applyLocal(command, next, metadata);
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
      await this.collect(this.allocations);
      throw error;
    }
  }
  private async replaceText(command: Command, next: Head, metadata: Metadata) {
    const id = String(command.payload.sheetId);
    for (const a of Object.keys(next.blocks))
      if (a.startsWith(`${id}/`)) {
        this.check();
        const cells = structuredClone((await this.block(a, next)).cells);
        const snapshot = {
          ...metadata.snapshot,
          sheets: metadata.snapshot.sheets.map((s) => (s.id === id ? { ...s, cells } : s)),
        };
        const before = JSON.stringify(cells);
        reduceCommand(snapshot, command);
        if (before !== JSON.stringify(cells)) await this.storeBlock(a, cells, next, metadata);
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
    metadata.snapshot = transformed;
    for (const [address, pointer] of Object.entries(next.blocks)) {
      this.check();
      const [sheetId, br, bc] = address.split('/');
      const removedSheet = command.type === 'core.sheet.remove' && sheetId === sourceSheet.id;
      const deletion =
        sheetId === sourceSheet.id && deletedBuckets.has(payload.axis === 'row' ? +br : +bc);
      const formula = pointer.formulas?.some(
        (name) => name === '*' || name.toLowerCase() === sourceSheet.name.toLowerCase(),
      );
      if (removedSheet) {
        this.changed.add(address);
        delete next.blocks[address];
        continue;
      }
      if (!deletion && !formula) continue;
      const working = {
        ...original,
        sheets: original.sheets.map((sh) => ({
          ...structuredClone(sh),
          cells: sh.id === sheetId ? structuredClone({} as Record<string, CellRecord>) : {},
        })),
      };
      working.sheets.find((sh) => sh.id === sheetId)!.cells = structuredClone(
        (await this.block(address, next)).cells,
      );
      reduceCommand(working, command);
      const output = working.sheets.find((sh) => sh.id === sheetId);
      if (output) await this.storeBlock(address, output.cells, next, metadata);
    }
  }
  private async sort(command: Command, next: Head, metadata: Metadata) {
    const p = command.payload as any,
      sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId)!;
    checkBounds(sh, p.range);
    // Validate range/merge constraints using the existing semantic reducer.
    assert(
      p.column >= p.range.startColumn && p.column < p.range.endColumn,
      'INVALID_RANGE',
      'Sort column outside range',
    );
    assert(
      !sh.merges.some((m) => intersects(m, p.range)),
      'INVALID_RANGE',
      'Unmerge before sorting',
    );
    const source = { ...next, blocks: { ...next.blocks } };
    for (const block of Object.values(source.blocks)) this.pins.add(block.id);
    let runs: Array<{ id: string; count: number; bytes: number }> = [];
    let scratchBytes = 0;
    type SortRow = { row: number; cells: Array<CellRecord | null>; value: CellValue };
    const compare = (a: SortRow, b: SortRow) => {
      const x = a.value,
        y = b.value;
      const cmp =
        x === y
          ? 0
          : x === null
            ? 1
            : y === null
              ? -1
              : typeof x === 'number' && typeof y === 'number'
                ? x - y
                : String(x).localeCompare(String(y), 'en');
      return (p.direction === 'asc' ? cmp : -cmp) || a.row - b.row;
    };
    try {
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
              (await this.block(blockAddress(sh.id, sh.rowOrder[r], sh.columnOrder[c]), source))
                .cells[key(sh.rowOrder[r], sh.columnOrder[c])] ?? null;
            workingBytes += encodedBytes(cell) * 4;
            assert(
              workingBytes <= 32 * 1024 * 1024,
              'LIMIT_EXCEEDED',
              'Sort run exceeds the operation working-set budget',
            );
            cells.push(cell);
          }
          const input = cells[p.column - p.range.startColumn]?.input;
          assert(
            input?.type !== 'formula',
            'UNSUPPORTED_FEATURE',
            'Sorting on formula results is not supported yet',
          );
          batch.push({ row: r, cells, value: input && 'value' in input ? input.value : null });
        }
        assert(
          encodedBytes(batch) * 4 <= 32 * 1024 * 1024,
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
        runs.push({ id: run, count: batch.length, bytes: runBytes });
      }
      // Pairwise merge passes keep only two rows in memory, independently of run count.
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
          let a = 0,
            b = 0,
            output = 0,
            bytes = 0;
          let x = await this.db.get<SortRow>('scratch', `${left.id}/00000000`),
            y = await this.db.get<SortRow>('scratch', `${right.id}/00000000`);
          while (x || y) {
            this.check();
            const takeLeft = !!x && (!y || compare(x, y) <= 0),
              value = takeLeft ? x! : y!;
            const record: [IDBValidKey, unknown] = [
              `${id}/${String(output++).padStart(8, '0')}`,
              value,
            ];
            const size = encodedBytes([record]);
            scratchBytes += size;
            bytes += size;
            this.stagingBytes = scratchBytes;
            assert(
              this.usedBytes + scratchBytes <= next.budget,
              'STORAGE_BUDGET',
              'Sort merge exceeds the storage budget',
            );
            await this.db.write('scratch', [record]);
            if (takeLeft)
              x =
                ++a < left.count
                  ? await this.db.get<SortRow>(
                      'scratch',
                      `${left.id}/${String(a).padStart(8, '0')}`,
                    )
                  : undefined;
            else
              y =
                ++b < right.count
                  ? await this.db.get<SortRow>(
                      'scratch',
                      `${right.id}/${String(b).padStart(8, '0')}`,
                    )
                  : undefined;
          }
          await this.db.write('scratch', [], [prefixRange(left.id), prefixRange(right.id)]);
          scratchBytes -= left.bytes + right.bytes;
          this.stagingBytes = scratchBytes;
          merged.push({ id, count: output, bytes });
        }
        runs = merged;
        pass++;
      }
      let position = 0;
      for (let row = p.range.startRow; row < p.range.endRow; row++) {
        this.check();
        const entry = (await this.db.get<SortRow>(
          'scratch',
          `${runs[0].id}/${String(position++).padStart(8, '0')}`,
        ))!;
        const grouped = new Map<string, Record<string, CellRecord>>();
        for (let c = p.range.startColumn; c < p.range.endColumn; c++) {
          const a = blockAddress(sh.id, sh.rowOrder[row], sh.columnOrder[c]);
          if (!grouped.has(a)) grouped.set(a, structuredClone((await this.block(a, next)).cells));
          const k = key(sh.rowOrder[row], sh.columnOrder[c]);
          const cell = entry.cells[c - p.range.startColumn];
          if (!cell) delete grouped.get(a)![k];
          else {
            const copied = structuredClone(cell);
            copied.rowId = sh.rowOrder[row];
            copied.columnId = sh.columnOrder[c];
            delete copied.cached;
            if (copied.input.type === 'formula')
              copied.input.expression = offsetFormula(copied.input.expression, row - entry.row, 0);
            grouped.get(a)![k] = copied;
          }
        }
        for (const [a, cells] of grouped) await this.storeBlock(a, cells, next, metadata);
      }
    } finally {
      await this.db.write('scratch', [], [prefixRange(this.id)]);
      this.stagingBytes = 0;
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
    const h = this.histories.get(id) ?? (await this.db.get<History>('records', id));
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
    const metadata =
      next.meta === this.head.meta ? this.metadata : await this.readMetadata(next.meta);
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
  async search(sheetId: string, query: string, after = -1) {
    const sh = this.metadata.snapshot.sheets.find((s) => s.id === sheetId)!;
    const q = query.toLowerCase();
    let first: { row: number; column: number } | undefined,
      next: typeof first,
      count = 0;
    const rows = new Map(sh.rowOrder.map((id, i) => [id, i])),
      cols = new Map(sh.columnOrder.map((id, i) => [id, i]));
    for (const [a] of Object.entries(this.head.blocks))
      if (a.startsWith(`${sheetId}/`))
        for (const cell of Object.values((await this.block(a)).cells)) {
          this.check();
          const row = rows.get(cell.rowId)!,
            column = cols.get(cell.columnId)!;
          const value = await this.calculator.value(sheetId, row, column);
          if (
            !formatValue(value, cell.numberFormat, this.metadata.snapshot.dateSystem)
              .toLowerCase()
              .includes(q)
          )
            continue;
          count++;
          const position = row * sh.columnOrder.length + column;
          if (!first || position < first.row * sh.columnOrder.length + first.column)
            first = { row, column };
          if (
            position > after &&
            (!next || position < next.row * sh.columnOrder.length + next.column)
          )
            next = { row, column };
        }
    return { match: next ?? first, count };
  }
  async clear() {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    await this.db.write('heads', [], [this.id]);
    await this.db.write('records', [], [prefixRange(this.id)]);
    await this.db.write('scratch', [], [prefixRange(this.id)]);
    this.cache.clear();
  }
  close() {
    this.release?.();
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
