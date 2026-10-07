import {
  applyPatches,
  createDraft,
  enablePatches,
  finishDraft,
  freeze,
  produceWithPatches,
  type Draft,
  type Patch,
} from 'immer';
import { FormulaCalculation } from './calculation.js';
import { reduceCommand, validateCommand } from './commands.js';
import { formatValue } from './format.js';
import { UndoHistory } from './history.js';
import { createSheet, createSnapshot, validateSnapshot } from './model.js';
import { Sheet } from './sheet.js';
import {
  assert,
  key,
  LIMITS,
  scalarValue,
  uid,
  type CellRecord,
  type CellStyle,
  type CellValue,
  type Command,
  type CommandEnvelope,
  type Commit,
  type SheetSnapshot,
  type WorkbookSnapshot,
} from './types.js';

enablePatches();
export class Workbook {
  private state: WorkbookSnapshot;
  private draft?: Draft<WorkbookSnapshot>;
  private batch: Command[] = [];
  private poisoned = false;
  private listeners = new Set<(commit: Commit) => void>();
  private publishing = false;
  private disposed = false;
  private readOnly = false;
  private undoHistory = new UndoHistory();
  private calculator = new FormulaCalculation(
    (id) => this.sheetData(id),
    (name) => this.getSheetByName(name)?.id,
  );
  private importRevision: number;
  private seen = new Map<string, { signature: string; commit: Commit }>();
  private cellCount: number;
  constructor(snapshot: WorkbookSnapshot = createSnapshot()) {
    this.state = freeze(validateSnapshot(snapshot), true);
    this.cellCount = this.state.sheets.reduce(
      (count, sheet) => count + Object.keys(sheet.cells).length,
      0,
    );
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
        this.calculator.clear();
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
      this.calculator.clear();
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
    const nextCellCount = this.countAfterPatches(next, forward);
    assert(nextCellCount <= LIMITS.cells, 'LIMIT_EXCEEDED', 'Cell limit reached');
    const entry = this.undoHistory.prepare(forward, inverse);
    const previousRevision = this.state.revision;
    this.state = next;
    this.cellCount = nextCellCount;
    this.calculator.clear();
    this.undoHistory.push(entry);
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
  private countAfterPatches(next: WorkbookSnapshot, patches: Patch[]): number {
    if (patches.some((patch) => patch.path[0] === 'sheets' && patch.path.length <= 2))
      return next.sheets.reduce((count, sheet) => count + Object.keys(sheet.cells).length, 0);
    let count = this.cellCount;
    for (const patch of patches) {
      if (patch.path[0] === 'sheets' && patch.path[2] === 'cells' && patch.path.length === 4) {
        if (patch.op === 'add') count++;
        if (patch.op === 'remove') count--;
      }
    }
    return count;
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
    return this.undoHistory.canUndo;
  }
  get canRedo() {
    return this.undoHistory.canRedo;
  }
  private history(redo: boolean) {
    this.assertWritable();
    assert(!this.draft, 'INVALID_ARGUMENT', 'Cannot undo inside a transaction');
    const entry = this.undoHistory.take(redo);
    if (!entry) return false;
    const previousRevision = this.revision;
    const changed = applyPatches(this.state, redo ? entry.forward : entry.inverse);
    this.state = produceWithPatches(changed, (d) => {
      d.revision = previousRevision + 1;
    })[0];
    this.cellCount = this.state.sheets.reduce(
      (count, sheet) => count + Object.keys(sheet.cells).length,
      0,
    );
    this.undoHistory.accept(entry, redo);
    this.calculator.clear();
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
    return this.calculator.value(sheetId, row, column, stack);
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
    this.calculator.clear();
    this.undoHistory.clear();
    this.seen.clear();
  }
}
export function createWorkbook(options: Parameters<typeof createSnapshot>[0] = {}) {
  return new Workbook(createSnapshot(options));
}
