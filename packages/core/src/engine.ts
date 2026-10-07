import { intersects } from './address.js';
import { parseFormula, offsetFormula, type AST } from '@opensheetjs/formula';
import { decodeBlock, encodeBlock, type EncodedBlock } from './block-codec.js';
import { AsyncCalculation } from './async-calculation.js';
import { reduceCommand, validateCommand } from './commands.js';
import { formatValue } from './format.js';
import { checkBounds, validateSnapshot } from './model.js';
import { BLOCK_ROWS, BLOCK_COLUMNS, ByteCache, Database, encodedBytes } from './storage.js';
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
};
type Head = {
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
type History = {
  before: string;
  after: string;
  changes: Array<[string, Pointer | null, Pointer | null]>;
  bytes: number;
  ranges: Selection[];
};
export type EngineView = {
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
  private indices = new WeakMap<object, { rows: Map<string, number>; cols: Map<string, number> }>();
  private sizes = new Map<string, number>();
  private cache: ByteCache<Block>;
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
        return (await this.block(blockAddress(sheet, row, column))).dependencies[
          key(sh.rowOrder[row], sh.columnOrder[column])
        ] as AST | null | undefined;
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
  async open(snapshot?: WorkbookSnapshot, budget = 256 * 1024 * 1024) {
    assert(
      Number.isSafeInteger(budget) && budget > 0,
      'INVALID_ARGUMENT',
      'Storage budget must be a positive integer',
    );
    await this.lock();
    const existing = await this.db.get<Head>('heads', this.id);
    if (existing) {
      this.head = existing;
      this.metadata = (await this.db.get<Metadata>('records', this.head.meta))!;
      assert(this.metadata, 'CORRUPT_STORAGE', 'Workbook metadata is missing');
      if (this.writer) await this.recover();
      else this.usedBytes = await this.liveBytes();
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
      this.usedBytes = encodedBytes([this.id, this.head]);
      this.metadata = {
        snapshot: { ...snapshot, sheets: snapshot.sheets.map((s) => ({ ...s, cells: {} })) },
        filtered: {},
      };
      await this.replace(snapshot, budget);
    }
    return this.view();
  }
  private async liveIds(head = this.head) {
    const ids = new Set([
      head.meta,
      ...Object.values(head.blocks).map((p) => p.id),
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
    let total = encodedBytes([this.id, this.head]);
    for (const id of await this.liveIds()) {
      const bytes = await this.db.get<number>('sizes', id);
      if (bytes !== undefined) total += bytes + encodedBytes([id, bytes]);
    }
    return total;
  }
  private async recover() {
    const live = await this.liveIds();
    this.usedBytes = encodedBytes([this.id, this.head]);
    this.sizes.clear();
    for await (const batch of this.db.scan<number>('sizes', prefixRange(this.id))) {
      const dead: string[] = [];
      for (const [k, v] of batch) {
        const id = String(k);
        if (!live.has(id)) dead.push(id);
        else {
          const bytes = v + encodedBytes([id, v]);
          this.usedBytes += bytes;
          this.sizes.set(id, bytes);
        }
      }
      if (dead.length) await this.db.write('records', [], dead);
    }
    await this.db.write('scratch', [], [prefixRange(this.id)]);
  }
  async reload() {
    this.head = (await this.db.get<Head>('heads', this.id))!;
    this.metadata = (await this.db.get<Metadata>('records', this.head.meta))!;
    this.cache.clear();
    this.calculator.clear();
    this.usedBytes = await this.liveBytes();
    return this.view();
  }
  private async allocate(value: unknown, budget = this.head.budget) {
    this.check();
    const id = `${this.id}/${uid('record')}`;
    const rawBytes = encodedBytes([id, value]);
    const bytes = rawBytes + encodedBytes([id, rawBytes]);
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
  private async collect(candidates: Iterable<string>) {
    const live = await this.liveIds();
    const dead = [...new Set(candidates)].filter((id) => !live.has(id));
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
  private async publish(next: Head, metadata: Metadata) {
    this.check();
    const bytes =
      this.usedBytes - encodedBytes([this.id, this.head]) + encodedBytes([this.id, next]);
    assert(bytes <= next.budget, 'STORAGE_BUDGET', 'Application storage budget exceeded');
    await this.db.write('heads', [[this.id, next]]);
    this.head = next;
    this.metadata = metadata;
    this.usedBytes = bytes;
    this.calculator.clear();
  }
  view(): EngineView {
    const used: Record<string, Rect> = {};
    for (const s of this.metadata.snapshot.sheets) {
      let endRow = 1,
        endColumn = 1;
      for (const [address, p] of Object.entries(this.head.blocks))
        if (address.startsWith(`${s.id}/`)) {
          endRow = Math.max(endRow, p.endRow);
          endColumn = Math.max(endColumn, p.endColumn);
        }
      for (const m of s.merges) {
        endRow = Math.max(endRow, m.endRow);
        endColumn = Math.max(endColumn, m.endColumn);
      }
      used[s.id] = { startRow: 0, startColumn: 0, endRow, endColumn };
    }
    return {
      snapshot: { ...this.metadata.snapshot, revision: this.head.revision },
      filtered: this.metadata.filtered,
      used,
      canUndo: !!this.head.past.length,
      canRedo: !!this.head.future.length,
      bytes: this.usedBytes,
      cacheBytes: this.cache.bytes + this.calculator.cacheBytes + this.seen.bytes,
      budget: this.head.budget,
    };
  }
  private async block(address: string, head = this.head): Promise<Block> {
    const p = head.blocks[address];
    if (!p) return { cells: {}, dependencies: {} };
    const cached = this.cache.get(p.id);
    if (cached) return cached;
    const encoded = await this.db.get<EncodedBlock>('records', p.id);
    assert(encoded, 'CORRUPT_STORAGE', 'Data block is missing');
    const block = decodeBlock(encoded);
    this.cache.set(p.id, block);
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
      return { rowId: '', columnId: '', input: { type: 'error', code: '#REF!' } };
    return (await this.block(blockAddress(sheet, row, column))).cells[
      key(sh.rowOrder[row], sh.columnOrder[column])
    ];
  }
  async read(sheetId: string, range: Rect, values = true) {
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
        rows: new Map(sh.rowOrder.map((id, i) => [id, i])),
        cols: new Map(sh.columnOrder.map((id, i) => [id, i])),
      };
      this.indices.set(sh, index);
    }
    const { rows, cols } = index;
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
      sh.rowOrder.slice(+br * BLOCK_ROWS, (+br + 1) * BLOCK_ROWS),
      sh.columnOrder.slice(+bc * BLOCK_COLUMNS, (+bc + 1) * BLOCK_COLUMNS),
      dependencies,
    );
    const id = await this.allocate(encoded, head.budget);
    head.blocks[address] = { id, bytes: this.sizes.get(id)!, count, endRow, endColumn };
    this.cache.set(id, block);
    await discardPrevious();
  }
  async replace(snapshot: WorkbookSnapshot, budget = this.head.budget) {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.cancelled = false;
    this.allocations = [];
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
          const a = blockAddress(sh.id, r, col);
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
      next.meta = await this.allocate(metadata, budget);
      await this.publish(next, metadata);
      await this.cleanup([...old, ...this.allocations]);
      return this.view();
    } catch (error) {
      await this.collect(this.allocations);
      throw error;
    }
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
    this.allocated.clear();
    this.pins.clear();
    const metadata: Metadata = structuredClone(this.metadata);
    const sh = metadata.snapshot.sheets[0];
    sh.rowOrder = Array.from({ length: Math.ceil(count / 100) }, (_, i) => `${sh.id}_r${i}`);
    sh.columnOrder = Array.from({ length: 100 }, (_, i) => `${sh.id}_c${i}`);
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
      next.meta = await this.allocate(metadata, budget);
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
    const out: string[] = [];
    for (
      let r = Math.floor(range.startRow / BLOCK_ROWS);
      r < Math.ceil(range.endRow / BLOCK_ROWS);
      r++
    )
      for (
        let c = Math.floor(range.startColumn / BLOCK_COLUMNS);
        c < Math.ceil(range.endColumn / BLOCK_COLUMNS);
        c++
      )
        out.push(`${sheetId}/${r}/${c}`);
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
      if (p.cells) for (const c of p.cells) addresses.add(blockAddress(sh.id, c.row, c.column));
      for (const range of [p.range, p.target])
        if (range) {
          checkBounds(sh, range);
          for (const a of this.addresses(sh.id, range)) addresses.add(a);
        }
    }
    let bytes = 0;
    const snapshot = structuredClone(metadata.snapshot);
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
      for (
        let r = +br * BLOCK_ROWS;
        r < Math.min((+br + 1) * BLOCK_ROWS, target!.rowOrder.length);
        r++
      )
        for (
          let c = +bc * BLOCK_COLUMNS;
          c < Math.min((+bc + 1) * BLOCK_COLUMNS, target!.columnOrder.length);
          c++
        ) {
          const k = key(target!.rowOrder[r], target!.columnOrder[c]);
          if (target!.cells[k]) cells[k] = target!.cells[k];
        }
      await this.storeBlock(address, cells, head, metadata);
    }
    for (const s of snapshot.sheets) s.cells = {};
    metadata.snapshot = snapshot;
  }
  async execute(inputs: Array<Command | CommandEnvelope>, label = 'edit') {
    assert(this.writer, 'READ_ONLY', 'Workbook is open in another tab');
    this.cancelled = false;
    this.allocations = [];
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
      metadata = structuredClone(this.metadata),
      next = structuredClone(old);
    const candidates = await this.liveIds();
    const ranges: Selection[] = [];
    try {
      for (const command of commands) {
        this.check();
        const p = command.payload as any;
        if (p.sheetId) {
          const sh = metadata.snapshot.sheets.find((s) => s.id === p.sheetId)!;
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
      next.meta =
        JSON.stringify(metadata) === JSON.stringify(this.metadata)
          ? old.meta
          : await this.allocate(metadata);
      const changes: History['changes'] = [];
      for (const a of new Set([...Object.keys(old.blocks), ...Object.keys(next.blocks)]))
        if (old.blocks[a]?.id !== next.blocks[a]?.id)
          changes.push([a, old.blocks[a] ?? null, next.blocks[a] ?? null]);
      const bytes =
        changes.reduce((n, [, a, b]) => n + (a?.bytes ?? 0) + (b?.bytes ?? 0), 0) +
        (next.meta === old.meta ? 0 : encodedBytes(metadata) + encodedBytes(this.metadata));
      assert(
        bytes <= 32 * 1024 * 1024,
        'LIMIT_EXCEEDED',
        'Change exceeds the 32 MiB undo budget; import as a new workbook',
      );
      const history: History = { before: old.meta, after: next.meta, changes, bytes, ranges };
      next.past.push(await this.allocate(history));
      next.future = [];
      next.historyBytes = bytes;
      for (const id of old.past)
        next.historyBytes += (await this.db.get<History>('records', id))!.bytes;
      while (next.past.length > 100 || next.historyBytes > 32 * 1024 * 1024) {
        const expired = next.past.shift()!;
        next.historyBytes -= (await this.db.get<History>('records', expired))!.bytes;
      }
      await this.publish(next, metadata);
      await this.cleanup([...candidates, ...this.allocations]);
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
    const original = structuredClone(metadata.snapshot);
    const sourceHead = structuredClone(next);
    for (const p of Object.values(sourceHead.blocks)) this.pins.add(p.id);
    const transformed = structuredClone(original);
    reduceCommand(transformed, command);
    next.blocks = {};
    const rows = new Map<string, Map<string, number>>(),
      cols = new Map<string, Map<string, number>>();
    for (const s of transformed.sheets) {
      rows.set(s.id, new Map(s.rowOrder.map((id, i) => [id, i])));
      cols.set(s.id, new Map(s.columnOrder.map((id, i) => [id, i])));
    }
    metadata.snapshot = transformed;
    for (const [address] of Object.entries(sourceHead.blocks)) {
      this.check();
      const sheetId = address.split('/')[0];
      const working = structuredClone(original);
      const source = working.sheets.find((s) => s.id === sheetId)!;
      source.cells = structuredClone((await this.block(address, sourceHead)).cells);
      reduceCommand(working, command);
      const output = working.sheets.find((s) => s.id === sheetId);
      if (!output) continue;
      const grouped = new Map<string, Record<string, CellRecord>>();
      for (const [k, cell] of Object.entries(output.cells)) {
        const r = rows.get(sheetId)!.get(cell.rowId)!,
          c = cols.get(sheetId)!.get(cell.columnId)!;
        const a = blockAddress(sheetId, r, c);
        if (!grouped.has(a)) grouped.set(a, structuredClone((await this.block(a, next)).cells));
        grouped.get(a)![k] = cell;
      }
      for (const [a, cells] of grouped) await this.storeBlock(a, cells, next, metadata);
      this.progress({
        phase: 'structure',
        completed: Object.keys(next.blocks).length,
        total: Object.keys(sourceHead.blocks).length,
      });
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
    const source = structuredClone(next);
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
              (await this.block(blockAddress(sh.id, r, c), source)).cells[
                key(sh.rowOrder[r], sh.columnOrder[c])
              ] ?? null;
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
          const a = blockAddress(sh.id, row, c);
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
    const next = structuredClone(this.head),
      from = redo ? next.future : next.past,
      to = redo ? next.past : next.future;
    const id = from.pop();
    if (!id) return { view: this.view(), commit: undefined };
    const h = (await this.db.get<History>('records', id))!;
    for (const [a, before, after] of h.changes) {
      const p = redo ? after : before;
      if (p) next.blocks[a] = p;
      else delete next.blocks[a];
    }
    next.meta = redo ? h.after : h.before;
    next.revision++;
    to.push(id);
    const metadata = (await this.db.get<Metadata>('records', next.meta))!;
    await this.publish(next, metadata);
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
