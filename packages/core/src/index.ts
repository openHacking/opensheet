import {
  freeze,
  enablePatches,
  produceWithPatches,
  applyPatches,
  createDraft,
  finishDraft,
  type Patch,
  type Draft,
} from 'immer';
import { parseFormula, evaluateFormula, type Reference } from '@opensheetjs/formula';
import {
  assert,
  OpenSheetError,
  key,
  uid,
  toInput,
  scalarValue,
  type WorkbookSnapshot,
  type CellRecord,
  type CellStyle,
  type CellInput,
  type CellValue,
  type Primitive,
  type Rect,
  type Commit,
  type Command,
  type CommandEnvelope,
  type SheetSnapshot,
} from './types.js';
import { parseRange, address, rangeAddress } from './address.js';
import { createSnapshot, createSheet, validateSnapshot, checkBounds } from './model.js';
import { commandSchemas, validateCommand, reduceCommand } from './commands.js';
export * from './types.js';
export * from './address.js';
export * from './model.js';
export { commandSchemas };
enablePatches();
type History = { forward: Patch[]; inverse: Patch[]; bytes: number };
export class Workbook {
  private state: WorkbookSnapshot;
  private draft?: Draft<WorkbookSnapshot>;
  private batch: Command[] = [];
  private poisoned = false;
  private listeners = new Set<(commit: Commit) => void>();
  private publishing = false;
  private disposed = false;
  private readOnly = false;
  private past: History[] = [];
  private future: History[] = [];
  private historyBytes = 0;
  private cache = new Map<string, CellValue>();
  private importRevision: number;
  private seen = new Map<string, { signature: string; commit: Commit }>();
  constructor(snapshot: WorkbookSnapshot = createSnapshot()) {
    this.state = freeze(validateSnapshot(snapshot), true);
    this.readOnly = this.state.extensions['opensheet.readOnly'] === true;
    this.importRevision = this.state.revision;
  }
  get revision() {
    this.assertActive();
    return this.current.revision;
  }
  get id() {
    this.assertActive();
    return this.current.workbookId;
  }
  private get current() {
    return this.draft ?? this.state;
  }
  private assertActive() {
    assert(!this.disposed, 'DISPOSED', 'Workbook has been disposed');
  }
  private assertWritable() {
    this.assertActive();
    assert(this.revision < Number.MAX_SAFE_INTEGER, 'LIMIT_EXCEEDED', 'Revision limit reached');
    assert(!this.readOnly, 'READ_ONLY', 'Workbook is read only');
    assert(!this.publishing, 'INVALID_ARGUMENT', 'Schedule changes outside event callbacks');
  }
  setReadOnly(value: boolean) {
    this.assertActive();
    this.readOnly = value || this.state.extensions['opensheet.readOnly'] === true;
  }
  onCommit(fn: (commit: Commit) => void) {
    this.assertActive();
    this.listeners.add(fn);
    return () => {
      this.listeners.delete(fn);
    };
  }
  toJSON(): WorkbookSnapshot {
    this.assertActive();
    const result = JSON.parse(JSON.stringify(this.current)) as WorkbookSnapshot;
    if (this.revision !== this.importRevision)
      for (const sh of result.sheets) for (const c of Object.values(sh.cells)) delete c.cached;
    return result;
  }
  getSheets() {
    this.assertActive();
    return this.current.sheetOrder.map((id) => new Sheet(this, id));
  }
  getSheetById(id: string) {
    this.assertActive();
    return this.current.sheets.some((s) => s.id === id) ? new Sheet(this, id) : undefined;
  }
  getSheetByName(name: string) {
    this.assertActive();
    const s = this.current.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase());
    return s ? new Sheet(this, s.id) : undefined;
  }
  sheetData(id: string): Readonly<SheetSnapshot> {
    this.assertActive();
    const sh = this.current.sheets.find((s) => s.id === id);
    assert(sh, 'INVALID_ARGUMENT', 'Sheet no longer exists');
    return this.draft ? freeze(JSON.parse(JSON.stringify(sh)) as SheetSnapshot, true) : sh;
  }
  getStyle(id?: string): CellStyle {
    this.assertActive();
    return id ? { ...this.current.styles[id] } : {};
  }
  getCell(sheetId: string, row: number, column: number): CellRecord | undefined {
    const sh = this.sheetData(sheetId),
      c = sh.cells[key(sh.rowOrder[row], sh.columnOrder[column])];
    return c ? JSON.parse(JSON.stringify(c)) : undefined;
  }
  addSheet(name: string, options: { rows?: number; columns?: number } = {}) {
    const sh = createSheet(name, options.rows, options.columns);
    this.execute({ type: 'core.sheet.add', payload: { sheet: sh } });
    return this.getSheetById(sh.id)!;
  }
  removeSheet(id: string) {
    return this.execute({ type: 'core.sheet.remove', payload: { sheetId: id } });
  }
  renameSheet(id: string, name: string) {
    return this.execute({ type: 'core.sheet.rename', payload: { sheetId: id, name } });
  }
  execute(input: Command | CommandEnvelope): Commit | undefined {
    this.assertWritable();
    try {
      const command = validateCommand(input);
      let signature = '';
      if ('protocolVersion' in input) {
        assert(!this.draft, 'INVALID_ARGUMENT', 'Envelopes cannot be nested in transactions');
        assert(
          input.protocolVersion === 1 &&
            input.workbookId === this.id &&
            typeof input.commandId === 'string' &&
            input.commandId.length > 0 &&
            input.commandId.length <= 100,
          'INVALID_ARGUMENT',
          'Invalid envelope',
        );
        signature = JSON.stringify(input);
        const old = this.seen.get(input.commandId);
        if (old) {
          assert(
            old.signature === signature,
            'INVALID_ARGUMENT',
            'Command id was reused with a different payload',
          );
          return structuredClone(old.commit);
        }
        assert(input.baseRevision === this.revision, 'REVISION_CONFLICT', 'Workbook has changed');
      }
      if (this.draft) {
        reduceCommand(this.draft, command);
        this.batch.push(command);
        this.cache.clear();
        return;
      }
      const [next, patches, inverse] = produceWithPatches(this.state, (draft) => {
        reduceCommand(draft, command);
        draft.revision++;
      });
      const commit = this.install(
        next,
        patches,
        inverse,
        [command],
        'edit',
        'commandId' in input ? input.commandId : uid('cmd'),
      );
      if ('protocolVersion' in input) {
        this.seen.set(input.commandId, { signature, commit: structuredClone(commit) });
        if (this.seen.size > 1000) this.seen.delete(this.seen.keys().next().value!);
      }
      return commit;
    } catch (error) {
      if (this.draft) this.poisoned = true;
      throw error;
    }
  }
  transaction<T>(options: { label?: string }, fn: () => T): T {
    this.assertWritable();
    assert(!this.draft, 'INVALID_ARGUMENT', 'Nested transactions are not supported');
    assert(
      fn.constructor.name !== 'AsyncFunction',
      'INVALID_ARGUMENT',
      'Transactions must be synchronous',
    );
    this.draft = createDraft(this.state);
    this.batch = [];
    this.poisoned = false;
    try {
      const result = fn();
      assert(
        !(result && typeof (result as any).then === 'function'),
        'INVALID_ARGUMENT',
        'Transactions must not return a Promise',
      );
      assert(!this.poisoned, 'INVALID_ARGUMENT', 'Transaction contains a failed command');
      if (!this.batch.length) return result;
      this.draft.revision++;
      let forward: Patch[] = [],
        inverse: Patch[] = [];
      const next = finishDraft(this.draft, (p, i) => {
        forward = p;
        inverse = i;
      });
      this.draft = undefined;
      this.install(next, forward, inverse, this.batch, options.label ?? 'transaction', uid('cmd'));
      return result;
    } finally {
      if (this.draft) {
        finishDraft(this.draft);
        this.draft = undefined;
      }
      this.batch = [];
      this.cache.clear();
    }
  }
  private install(
    next: WorkbookSnapshot,
    forward: Patch[],
    inverse: Patch[],
    commands: Command[],
    source: string,
    commandId: string,
  ): Commit {
    const bytes = JSON.stringify([forward, inverse]).length * 2;
    assert(
      bytes <= 32 * 1024 * 1024,
      'LIMIT_EXCEEDED',
      'Change exceeds the undo budget; import as a new workbook',
    );
    const previousRevision = this.state.revision;
    this.state = next;
    this.cache.clear();
    this.past.push({ forward, inverse, bytes });
    this.historyBytes += bytes;
    this.future = [];
    while (this.past.length > 100 || this.historyBytes > 32 * 1024 * 1024)
      this.historyBytes -= this.past.shift()!.bytes;
    const commit = {
      commandId,
      previousRevision,
      revision: next.revision,
      source,
      commands,
      changedRanges: [],
    };
    this.emit(commit);
    return commit;
  }
  private emit(commit: Commit) {
    this.publishing = true;
    try {
      for (const fn of this.listeners) {
        try {
          fn(structuredClone(commit));
        } catch (e) {
          console.error('OpenSheet listener failed', e);
        }
      }
    } finally {
      this.publishing = false;
    }
  }
  undo() {
    return this.history(false);
  }
  redo() {
    return this.history(true);
  }
  get canUndo() {
    return this.past.length > 0;
  }
  get canRedo() {
    return this.future.length > 0;
  }
  private history(redo: boolean) {
    this.assertWritable();
    assert(!this.draft, 'INVALID_ARGUMENT', 'Cannot undo inside a transaction');
    const from = redo ? this.future : this.past,
      to = redo ? this.past : this.future,
      entry = from.pop();
    if (!entry) return false;
    const previousRevision = this.revision;
    const changed = applyPatches(this.state, redo ? entry.forward : entry.inverse);
    this.state = produceWithPatches(changed, (d) => {
      d.revision = previousRevision + 1;
    })[0];
    to.push(entry);
    this.historyBytes += redo ? entry.bytes : -entry.bytes;
    this.cache.clear();
    this.emit({
      commandId: uid('cmd'),
      previousRevision,
      revision: this.revision,
      source: redo ? 'redo' : 'undo',
      commands: [],
      changedRanges: [],
    });
    return true;
  }
  value(sheetId: string, row: number, column: number, stack = new Set<string>()): CellValue {
    this.assertActive();
    const sh = this.sheetData(sheetId);
    if (row >= sh.rowOrder.length || column >= sh.columnOrder.length || row < 0 || column < 0)
      return { error: '#REF!' };
    const k = `${sheetId}:${row}:${column}`;
    if (this.cache.has(k)) return this.cache.get(k)!;
    if (stack.has(k)) return { error: '#CYCLE!' };
    if (stack.size >= 128) return { error: '#LIMIT!' };
    const c = sh.cells[key(sh.rowOrder[row], sh.columnOrder[column])];
    if (!c) return null;
    if (c.input.type !== 'formula') return scalarValue(c.input);
    stack.add(k);
    let result: CellValue;
    try {
      result = evaluateFormula(parseFormula(c.input.expression), (ref: Reference) => {
        const target = ref.sheet ? this.getSheetByName(ref.sheet)?.id : sheetId;
        return target ? this.value(target, ref.row, ref.column, stack) : { error: '#REF!' };
      });
    } catch (e) {
      result = { error: e instanceof Error ? e.message : '#VALUE!' };
    } finally {
      stack.delete(k);
    }
    this.cache.set(k, result);
    return result;
  }
  display(sheetId: string, row: number, column: number): string {
    const c = this.getCell(sheetId, row, column);
    let v = this.value(sheetId, row, column);
    if (v && typeof v === 'object') {
      if (c?.cached && this.revision === this.importRevision)
        return `${formatValue(scalarValue(c.cached), c.numberFormat, this.current.dateSystem)} †`;
      return v.error;
    }
    return formatValue(
      v,
      c?.numberFormat ?? this.getStyle(c?.styleId).numberFormat,
      this.current.dateSystem,
    );
  }
  calculation = {
    calculate: async (options?: { revision?: number }) => {
      this.assertActive();
      assert(
        !options || options.revision === undefined || options.revision === this.revision,
        'REVISION_CONFLICT',
        'Calculation revision changed',
      );
      for (const sh of this.getSheets()) {
        const data = this.sheetData(sh.id);
        const rows = new Map(data.rowOrder.map((id, i) => [id, i])),
          cols = new Map(data.columnOrder.map((id, i) => [id, i]));
        for (const c of Object.values(data.cells))
          if (c.input.type === 'formula')
            this.value(sh.id, rows.get(c.rowId)!, cols.get(c.columnId)!);
      }
      return { revision: this.revision };
    },
    getState: () => ({ revision: this.revision, state: 'ready' as const }),
  };
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.listeners.clear();
    this.cache.clear();
    this.past = [];
    this.future = [];
    this.seen.clear();
  }
}
export function formatValue(
  value: CellValue,
  format?: string,
  dateSystem: '1900' | '1904' = '1900',
): string {
  if (value === null) return '';
  if (typeof value === 'object') return value.error;
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  if (typeof value !== 'number' || !format || format === 'General') return String(value);
  if (/[yd]/i.test(format)) {
    if (dateSystem === '1900' && Math.floor(value) === 60) return '1900-02-29';
    const origin = dateSystem === '1904' ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 31);
    const day = value - (dateSystem === '1900' && value >= 60 ? 1 : 0);
    const date = new Date(origin + day * 86400000);
    return Number.isNaN(date.getTime()) ? String(value) : date.toISOString().slice(0, 10);
  }
  const decimals = Math.min(12, format.match(/\.([0#]+)/)?.[1].length ?? 0);
  let text = (format.includes('%') ? value * 100 : value).toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    useGrouping: format.includes(','),
  });
  if (format.includes('%')) text += '%';
  if (format.includes('$')) text = '$' + text;
  return text;
}
export class Sheet {
  constructor(
    public readonly workbook: Workbook,
    public readonly id: string,
  ) {}
  get name() {
    return this.workbook.sheetData(this.id).name;
  }
  get rowCount() {
    return this.workbook.sheetData(this.id).rowOrder.length;
  }
  get columnCount() {
    return this.workbook.sheetData(this.id).columnOrder.length;
  }
  range(input: string | Rect) {
    const r = parseRange(input);
    checkBounds(this.workbook.sheetData(this.id) as SheetSnapshot, r);
    return new Range(this, r);
  }
  getUsedRange() {
    const s = this.workbook.sheetData(this.id),
      rs = new Map(s.rowOrder.map((id, i) => [id, i])),
      cs = new Map(s.columnOrder.map((id, i) => [id, i]));
    let endRow = 1,
      endColumn = 1;
    for (const c of Object.values(s.cells)) {
      endRow = Math.max(endRow, rs.get(c.rowId)! + 1);
      endColumn = Math.max(endColumn, cs.get(c.columnId)! + 1);
    }
    for (const m of s.merges) {
      endRow = Math.max(endRow, m.endRow);
      endColumn = Math.max(endColumn, m.endColumn);
    }
    return this.range({ startRow: 0, startColumn: 0, endRow, endColumn });
  }
  insertRows(index: number, count = 1) {
    return this.axis('row', index, count, true);
  }
  deleteRows(index: number, count = 1) {
    return this.axis('row', index, count, false);
  }
  insertColumns(index: number, count = 1) {
    return this.axis('column', index, count, true);
  }
  deleteColumns(index: number, count = 1) {
    return this.axis('column', index, count, false);
  }
  private axis(axis: string, index: number, count: number, insert: boolean) {
    assert(
      Number.isInteger(count) && count > 0 && count <= 100000,
      'INVALID_ARGUMENT',
      'Invalid count',
    );
    return this.workbook.execute({
      type: insert ? 'core.axis.insert' : 'core.axis.delete',
      payload: {
        sheetId: this.id,
        axis,
        index,
        ...(insert ? { ids: Array.from({ length: count }, () => uid(axis)) } : { count }),
      },
    });
  }
  setRowHeight(index: number, size: number) {
    return this.meta('row', index, { size });
  }
  setColumnWidth(index: number, size: number) {
    return this.meta('column', index, { size });
  }
  setRowHidden(index: number, hidden: boolean) {
    return this.meta('row', index, { hidden });
  }
  setColumnHidden(index: number, hidden: boolean) {
    return this.meta('column', index, { hidden });
  }
  private meta(axis: string, index: number, props: object) {
    return this.workbook.execute({
      type: 'core.axis.meta',
      payload: { sheetId: this.id, axis, index, ...props },
    });
  }
  setFreeze(rows: number, columns = 0) {
    return this.workbook.execute({
      type: 'core.sheet.freeze',
      payload: { sheetId: this.id, rows, columns },
    });
  }
  setFilter(filter: { column: number; query: string } | null) {
    return this.workbook.execute({
      type: 'core.sheet.filter',
      payload: { sheetId: this.id, filter },
    });
  }
  sort(range: string | Rect, column: number, direction: 'asc' | 'desc' = 'asc') {
    return this.workbook.execute({
      type: 'core.sheet.sort',
      payload: { sheetId: this.id, range: parseRange(range), column, direction },
    });
  }
}
export class Range {
  constructor(
    public readonly sheet: Sheet,
    public readonly bounds: Rect,
  ) {}
  get address() {
    return rangeAddress(this.bounds);
  }
  private matrix<T>(fn: (r: number, c: number) => T): T[][] {
    const b = this.bounds;
    return Array.from({ length: b.endRow - b.startRow }, (_, i) =>
      Array.from({ length: b.endColumn - b.startColumn }, (_, j) =>
        fn(b.startRow + i, b.startColumn + j),
      ),
    );
  }
  getValues() {
    return this.matrix((r, c) => this.sheet.workbook.value(this.sheet.id, r, c));
  }
  getDisplayValues() {
    return this.matrix((r, c) => this.sheet.workbook.display(this.sheet.id, r, c));
  }
  getFormulas() {
    return this.matrix((r, c) => {
      const input = this.sheet.workbook.getCell(this.sheet.id, r, c)?.input;
      return input?.type === 'formula' ? input.expression : null;
    });
  }
  getInputs() {
    return this.matrix(
      (r, c) => this.sheet.workbook.getCell(this.sheet.id, r, c)?.input ?? { type: 'blank' },
    );
  }
  private write<T>(matrix: T[][], map: (value: T) => CellInput) {
    const b = this.bounds;
    assert(
      matrix.length === b.endRow - b.startRow &&
        matrix.every((row) => row.length === b.endColumn - b.startColumn),
      'INVALID_ARGUMENT',
      'Matrix dimensions must match the range',
    );
    return this.sheet.workbook.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: this.sheet.id,
        cells: matrix.flatMap((row, i) =>
          row.map((value, j) => ({
            row: b.startRow + i,
            column: b.startColumn + j,
            input: map(value),
          })),
        ),
      },
    });
  }
  setValues(values: Primitive[][]) {
    return this.write(values, toInput);
  }
  copyTo(target: string | Rect) {
    return this.command('core.cells.copy', { target: parseRange(target) });
  }
  fillDown() {
    return this.command('core.cells.fillDown');
  }
  setFormulas(values: string[][]) {
    return this.write(values, (expression) => ({
      type: 'formula',
      expression: expression.replace(/^=/, ''),
    }));
  }
  setInput(value: string) {
    assert(
      this.bounds.endRow - this.bounds.startRow === 1 &&
        this.bounds.endColumn - this.bounds.startColumn === 1,
      'INVALID_RANGE',
      'setInput expects one cell',
    );
    return this.write([[value]], parseInput);
  }
  setStyle(style: CellStyle) {
    return this.command('core.cells.style', { style });
  }
  clear(options: { all?: boolean } = {}) {
    return this.command('core.cells.clear', { all: options.all ?? false });
  }
  merge(options: { discardCoveredValues?: boolean } = {}) {
    return this.command('core.cells.merge', {
      discardCoveredValues: options.discardCoveredValues ?? false,
    });
  }
  unmerge() {
    return this.command('core.cells.unmerge');
  }
  private command(type: string, extra: object = {}) {
    return this.sheet.workbook.execute({
      type,
      payload: { sheetId: this.sheet.id, range: this.bounds, ...extra },
    });
  }
}
export function parseInput(text: string): CellInput {
  if (text.startsWith('=')) return { type: 'formula', expression: text.slice(1) };
  if (text.startsWith("'")) return { type: 'string', value: text.slice(1) };
  if (text === '') return { type: 'blank' };
  if (/^(true|false)$/i.test(text))
    return { type: 'boolean', value: text.toLowerCase() === 'true' };
  if (
    /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:e[+-]?\d+)?$/i.test(text) &&
    text.replace(/[^\d]/g, '').length <= 15
  )
    return toInput(Number(text));
  return { type: 'string', value: text };
}
export function createWorkbook(options: Parameters<typeof createSnapshot>[0] = {}) {
  return new Workbook(createSnapshot(options));
}
