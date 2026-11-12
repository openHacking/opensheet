import {
  address,
  rangeAddress,
  OpenSheetError,
  LIMITS,
  parseInput,
  type CellInput,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
import { canInsertFormulaReference } from './formula-references.js';
import type { GridLayout } from './layout.js';
import type { GridSelection } from './selection.js';
export class CellEditor {
  readonly element = document.createElement('textarea');
  readonly recovery = document.createElement('section');
  private pendingWrites = 0;
  isEditing = false;
  isComposing = false;
  private editPosition = { row: 0, column: 0 };
  private editIdentity: { rowId: number; columnId: number; input?: CellInput | null } = {
    rowId: 0,
    columnId: 0,
  };
  private editRevision = 0;
  private origin?: Selection;
  private reference?: { start: number; end: number };
  private generation = 0;
  private pending?: { generation: number; selection: Selection };
  private disposed = false;
  private baseline?: string;
  constructor(
    private book: Workbook,
    private getSheetId: () => string,
    private getSelection: () => Selection,
    private selection: GridSelection,
    private layout: GridLayout,
    private prepareLayout: () => void,
    private viewport: HTMLElement,
    private scroll: HTMLElement,
    private readOnly: () => boolean,
    private onError: (error: unknown) => void,
    private onEnd: () => void,
    private onDraft: (dirty: boolean) => void,
  ) {
    this.element.className = 'os-cell-editor';
    this.element.setAttribute('aria-label', 'Edit cell');
    this.element.hidden = true;
    this.recovery.className = 'os-draft-recovery';
    this.recovery.setAttribute('aria-label', 'Unconfirmed cell drafts');
    this.recovery.hidden = true;
  }
  async start(text?: string) {
    if (this.disposed || this.readOnly() || this.isEditing) return;
    const generation = ++this.generation;
    this.pending = undefined;
    try {
      const s = this.getSelection(),
        m = this.selection.mergeAt(s.startRow, s.startColumn);
      const focused = document.activeElement;
      const position = {
        row: m?.startRow ?? s.startRow,
        column: m?.startColumn ?? s.startColumn,
      };
      const sheet = this.book.sheetData(s.sheetId);
      const identity = {
        rowId: sheet.rowOrder[position.row],
        columnId: sheet.columnOrder[position.column],
      };
      const revision = this.book.revision;
      const loaded = this.book.isLoaded(s.sheetId, position.row, position.column);
      if (text === undefined) this.pending = { generation, selection: { ...s } };
      const c =
        text === undefined
          ? await this.book.getCell(s.sheetId, position.row, position.column)
          : loaded
            ? this.book.peekCell(s.sheetId, position.row, position.column)
            : undefined;
      if (
        generation !== this.generation ||
        this.disposed ||
        this.readOnly() ||
        document.activeElement !== focused
      ) {
        if (this.pending?.generation === generation) this.pending = undefined;
        return;
      }
      this.pending = undefined;
      this.origin = { ...s };
      this.editPosition = position;
      this.editIdentity =
        text === undefined || loaded
          ? { ...identity, input: structuredClone(c?.input ?? null) }
          : identity;
      this.editRevision = revision;
      this.element.value =
        text ??
        (c?.input.type === 'formula'
          ? '=' + c.input.expression
          : c?.input.type === 'error'
            ? c.input.code
            : c?.input.type === 'blank'
              ? ''
              : String(c && 'value' in c.input ? c.input.value : ''));
      this.baseline = text === undefined ? this.element.value : undefined;
      this.isEditing = true;
      this.updateDraft();
      this.element.hidden = false;
      this.position();
      this.element.focus();
      if (text === undefined) this.element.select();
    } catch (error) {
      if (generation === this.generation && !this.disposed) {
        this.pending = undefined;
        this.onError(error);
      }
    }
  }
  position() {
    this.prepareLayout();
    const { row, column } = this.editPosition,
      p = this.layout.position(
        row,
        column,
        this.book.sheetData(this.getSheetId()).freeze,
        this.scroll,
      ),
      merge = this.selection.mergeAt(row, column),
      baseWidth =
        this.layout.columnOffsets[merge?.endColumn ?? column + 1] -
        this.layout.columnOffsets[column],
      baseHeight = this.layout.rowOffsets[merge?.endRow ?? row + 1] - this.layout.rowOffsets[row],
      margin = 4,
      availableWidth = Math.max(80, this.viewport.clientWidth - margin),
      availableHeight = Math.max(30, this.viewport.clientHeight - margin),
      width = Math.min(availableWidth, Math.max(100, baseWidth));
    this.element.style.width = `${width}px`;
    this.element.style.height = 'auto';
    const height = Math.min(availableHeight, Math.max(30, baseHeight, this.element.scrollHeight));
    Object.assign(this.element.style, {
      left: `${Math.max(0, Math.min(p.x, this.viewport.clientWidth - width))}px`,
      top: `${Math.max(0, Math.min(p.y, this.viewport.clientHeight - height))}px`,
      height: `${height}px`,
    });
  }
  finish(): boolean {
    if (!this.isEditing) return true;
    try {
      if (this.baseline !== undefined && this.element.value === this.baseline) {
        this.cancel();
        return true;
      }
      const state = this.book.persistenceState;
      if (state.state === 'failed' && !state.acknowledgeable) {
        throw new OpenSheetError(
          state.error?.code ?? 'WORKER_FAILED',
          state.error?.message ?? 'Reopen the workbook before saving this draft',
        );
      }
      if (this.pendingWrites + this.recovery.childElementCount >= 32)
        throw new OpenSheetError(
          'LIMIT_EXCEEDED',
          'Resolve unsaved drafts before editing more cells',
        );
      const value = this.element.value;
      if (value.length > LIMITS.text + 1)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Cell input is too long');
      const origin = { ...this.origin! };
      const cellAddress = address(this.editPosition.row, this.editPosition.column);
      const sheet = this.book.getSheetById(origin.sheetId)!;
      const location = `${sheet.name}!${cellAddress}`;
      const command = {
        type: 'core.cells.set',
        payload: {
          sheetId: origin.sheetId,
          cells: [{ ...this.editPosition, input: parseInput(value), expected: this.editIdentity }],
        },
      };
      const write = this.book.execute(
        this.editIdentity.input === undefined
          ? {
              ...command,
              protocolVersion: 1,
              commandId: crypto.randomUUID(),
              workbookId: this.book.id,
              baseRevision: this.editRevision,
            }
          : command,
      );
      this.pendingWrites++;
      void write.then(
        () => {
          this.pendingWrites--;
        },
        (error) => {
          this.pendingWrites--;
          if (this.disposed) return;
          const entry = document.createElement('div');
          const label = document.createElement('label');
          label.textContent = `Unconfirmed draft from ${location}`;
          const draft = document.createElement('textarea');
          draft.readOnly = true;
          draft.value = value;
          label.append(draft);
          const dismiss = document.createElement('button');
          dismiss.type = 'button';
          dismiss.textContent = 'Dismiss draft';
          dismiss.setAttribute('aria-label', `Dismiss draft from ${location}`);
          dismiss.onclick = () => {
            entry.remove();
            this.recovery.hidden = !this.recovery.childElementCount;
            this.updateDraft();
          };
          entry.append(label, dismiss);
          this.recovery.append(entry);
          this.recovery.hidden = false;
          this.updateDraft();
          this.onError(error);
        },
      );
      this.cancel();
      return true;
    } catch (e) {
      this.onError(e);
      this.element.focus();
      return false;
    }
  }
  cancel() {
    this.generation++;
    this.pending = undefined;
    this.isEditing = false;
    this.isComposing = false;
    this.updateDraft();
    this.element.hidden = true;
    this.reference = undefined;
    this.onEnd();
  }
  updateDraft() {
    this.onDraft(
      this.recovery.childElementCount > 0 ||
        (this.isEditing && (this.baseline === undefined || this.element.value !== this.baseline)),
    );
  }
  get isPending() {
    return !!this.pending;
  }
  cancelPending() {
    if (this.pending) {
      this.pending = undefined;
      this.generation++;
    }
  }
  selectionChanged(selection: Selection) {
    const pending = this.pending?.selection;
    if (
      pending &&
      (pending.sheetId !== selection.sheetId ||
        pending.startRow !== selection.startRow ||
        pending.endRow !== selection.endRow ||
        pending.startColumn !== selection.startColumn ||
        pending.endColumn !== selection.endColumn)
    )
      this.cancelPending();
  }
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.cancel();
  }
  getOrigin() {
    return this.origin ? { ...this.origin } : undefined;
  }
  takeFormulaDraft() {
    if (!this.canPickReference() || !this.origin) return undefined;
    const draft = {
      value: this.element.value,
      start: this.element.selectionStart,
      end: this.element.selectionEnd,
      origin: { ...this.origin },
      identity: { ...this.editIdentity },
      revision: this.editRevision,
    };
    this.cancel();
    return draft;
  }
  canPickReference() {
    return (
      this.isEditing &&
      !this.isComposing &&
      !this.readOnly() &&
      document.activeElement === this.element &&
      canInsertFormulaReference(this.element.value, this.element.selectionStart)
    );
  }
  beginReference() {
    if (!this.canPickReference()) return false;
    this.reference = { start: this.element.selectionStart, end: this.element.selectionEnd };
    return true;
  }
  updateReference(selection: Selection) {
    if (!this.reference || selection.sheetId !== this.origin?.sheetId) return;
    const text = rangeAddress(selection);
    this.element.setRangeText(text, this.reference.start, this.reference.end, 'end');
    this.reference.end = this.reference.start + text.length;
    this.element.dispatchEvent(new Event('input', { bubbles: true }));
  }
  endReference() {
    this.reference = undefined;
  }
}
