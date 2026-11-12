import { OpenSheetError, Workbook, type Selection } from '@opensheetjs/core';
import { copyText, pasteText, pasteHTML } from './clipboard.js';
import { copyClipboard } from './copy-clipboard.js';
import { CellEditor } from './editor.js';
import { FormulaAssist } from './formula-assist.js';
import { formulaReferenceColor, resolveFormulaReferenceSpans } from './formula-references.js';
import { LEFT, TOP, type CellPosition } from './geometry.js';
import { GridInput } from './input.js';
import { GridLayout } from './layout.js';
import { paintGrid } from './paint.js';
import { GridSelection } from './selection.js';
import { GridScrollbars } from './scrollbars.js';
import type { GridOptions } from './types.js';
export class CanvasGrid {
  readonly element: HTMLDivElement;
  readonly scroller: HTMLDivElement;
  private spacer: HTMLDivElement;
  private scrollbars: GridScrollbars;
  private canvas: HTMLCanvasElement;
  private cellEditor: CellEditor;
  private cellReferenceMirror = document.createElement('div');
  private formulaAssist: FormulaAssist;
  private input: GridInput;
  private ariaCell: HTMLDivElement;
  private sheetId: string;
  private selectionState: GridSelection;
  private get selected() {
    return this.selectionState.value;
  }
  private set selected(value: Selection) {
    this.selectionState.value = value;
  }
  private layoutState = new GridLayout();
  private raf = 0;
  private abort = new AbortController();
  private resize: ResizeObserver;
  private stop: () => void;
  private stopData: () => void;
  private stopDispose: () => void;
  private fetching?: Promise<void>;
  private fetchError?: unknown;
  private disposed = false;
  private selectionRevision = 0;
  private dataJumpAbort?: AbortController;
  private dataJumpEpoch = 0;
  private dataJumpProcessing?: number;
  private dataJumpQueue: Array<{ rows: number; columns: number; extend: boolean }> = [];
  private applyingDataJump = false;
  private keyboardSelection?: {
    anchor: { row: number; column: number };
    focus: { row: number; column: number };
  };
  private zoom = 1;
  private formulaReferences: Selection[] = [];
  constructor(
    private container: HTMLElement,
    private book: Workbook,
    private options: GridOptions = {},
  ) {
    this.sheetId = book.getSheets().find((s) => !book.sheetData(s.id).hidden)!.id;
    this.selectionState = new GridSelection(
      { sheetId: this.sheetId, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 },
      (id) => book.sheetData(id).merges,
    );
    this.element = document.createElement('div');
    this.element.className = 'os-grid';
    this.element.tabIndex = 0;
    this.element.setAttribute('role', 'grid');
    this.element.setAttribute('aria-label', 'Spreadsheet');
    this.scroller = document.createElement('div');
    this.scroller.className = 'os-scroll';
    this.scroller.id = `os-scroll-${crypto.randomUUID()}`;
    this.spacer = document.createElement('div');
    this.scroller.append(this.spacer);
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'os-canvas';
    this.canvas.setAttribute('aria-hidden', 'true');
    this.cellEditor = new CellEditor(
      book,
      () => this.sheetId,
      () => this.selection,
      this.selectionState,
      this.layoutState,
      () => {
        if (this.layoutState.dirty) this.layout();
      },
      this.scroller,
      this.scroller,
      () => !!this.options.readOnly,
      (error) => this.options.onError?.(error),
      () => {
        this.setFormulaReferences([]);
        this.cellReferenceMirror.hidden = true;
        this.cellReferenceMirror.replaceChildren();
      },
      (dirty) => this.options.onDraft?.(dirty),
    );
    this.formulaAssist = new FormulaAssist(this.cellEditor.element, this.abort.signal, 'cell');
    this.cellReferenceMirror.className = 'os-cell-reference-mirror';
    this.cellReferenceMirror.setAttribute('aria-hidden', 'true');
    this.cellReferenceMirror.hidden = true;
    this.ariaCell = document.createElement('div');
    this.ariaCell.id = `os-active-${crypto.randomUUID()}`;
    const ariaRow = document.createElement('div');
    ariaRow.className = 'os-sr-only';
    ariaRow.setAttribute('role', 'row');
    ariaRow.append(this.ariaCell);
    this.ariaCell.setAttribute('role', 'gridcell');
    this.element.setAttribute('aria-activedescendant', this.ariaCell.id);
    this.element.append(
      this.scroller,
      this.canvas,
      this.cellEditor.element,
      this.cellEditor.recovery,
      this.cellReferenceMirror,
      this.formulaAssist.element,
      ariaRow,
    );
    this.element.setAttribute('aria-readonly', String(!!this.options.readOnly));
    container.append(this.element);
    const signal = this.abort.signal;
    this.scrollbars = new GridScrollbars(this.element, this.scroller, signal);
    this.scroller.addEventListener(
      'scroll',
      () => {
        if (this.cellEditor.isEditing) this.positionEditor();
        this.schedule();
      },
      { signal },
    );
    this.input = new GridInput(this.element, book, this.cellEditor, {
      selection: () => this.selection,
      select: (s) => this.setSelection(s),
      selectBetween: (a, b) => this.selectBetween(a, b),
      hit: (e) => this.hit(e),
      move: (r, c, extend) => this.move(r, c, extend),
      page: (direction, extend) => this.move(direction, 0, extend, true),
      dataJump: (rows, columns, extend) => this.run(() => this.dataJump(rows, columns, extend)),
      cancelDataJump: () => this.cancelDataJump(),
      end: (extend) => {
        const used = this.book.getSheetById(this.sheetId)!.getUsedRange().bounds;
        this.move(0, 0, extend, false, { row: used.endRow - 1, column: used.endColumn - 1 });
      },
      focus: () => this.focus(),
      run: (fn) => this.run(fn),
      range: () => this.activeRange(),
      readOnly: () => !!this.options.readOnly,
      paste: (text, html) => (html ? this.pasteHTML(html, text) : this.pasteText(text)),
      positionEditor: () => this.positionEditor(),
      startKeyboardReference: () => this.options.onKeyboardReference?.() ?? false,
      referencePicker: options.referencePicker,
    });
    const refreshCellReferences = () => {
      if (this.disposed || !this.cellEditor.isEditing || this.cellEditor.isComposing) return;
      const source = this.cellEditor.element.value;
      const references = resolveFormulaReferenceSpans(this.book, this.sheetId, source);
      this.setFormulaReferences(references.map(({ selection }) => selection));
      const nodes: Node[] = [];
      let position = 0;
      references.forEach(({ start, end }, index) => {
        nodes.push(document.createTextNode(source.slice(position, start)));
        const mark = document.createElement('b');
        mark.textContent = source.slice(start, end);
        mark.style.textDecorationColor = formulaReferenceColor(index);
        nodes.push(mark);
        position = end;
      });
      nodes.push(document.createTextNode(source.slice(position)));
      this.cellReferenceMirror.replaceChildren(...nodes);
      this.cellReferenceMirror.hidden = references.length === 0;
      this.syncCellReferenceMirror();
    };
    this.cellEditor.element.addEventListener(
      'focus',
      () => {
        this.positionEditor();
        refreshCellReferences();
      },
      { signal },
    );
    this.cellEditor.element.addEventListener(
      'input',
      () => {
        this.cellEditor.updateDraft();
        refreshCellReferences();
      },
      { signal },
    );
    this.cellEditor.element.addEventListener(
      'compositionstart',
      () => {
        this.cellReferenceMirror.hidden = true;
      },
      { signal },
    );
    this.cellEditor.element.addEventListener('scroll', () => this.syncCellReferenceMirror(), {
      signal,
    });
    this.cellEditor.element.addEventListener(
      'compositionend',
      () => {
        queueMicrotask(refreshCellReferences);
      },
      { signal },
    );
    this.stop = book.onCommit((commit) => {
      this.keyboardSelection = undefined;
      if (!book.getSheetById(this.sheetId))
        this.setSheet(book.getSheets().find((s) => !book.sheetData(s.id).hidden)!.id);
      const sh = book.sheetData(this.sheetId);
      if (
        !commit.commands.length ||
        sh.filter ||
        commit.commands.some((command) =>
          [
            'core.axis.insert',
            'core.axis.delete',
            'core.axis.meta',
            'core.axis.visibility',
            'core.axis.visibilityAll',
            'core.axis.resize',
            'core.sheet.freeze',
            'core.sheet.filter',
          ].includes(command.type),
        )
      )
        this.layoutState.dirty = true;
      const sheet = book.getSheetById(this.sheetId)!;
      const previousSelection = this.selected;
      this.selected = {
        ...this.selected,
        startRow: Math.min(this.selected.startRow, sheet.rowCount - 1),
        endRow: Math.min(this.selected.endRow, sheet.rowCount),
        startColumn: Math.min(this.selected.startColumn, sheet.columnCount - 1),
        endColumn: Math.min(this.selected.endColumn, sheet.columnCount),
      };
      this.selected = this.expandSelection(this.selected);
      this.cellEditor.selectionChanged(this.selected);
      if (
        previousSelection.sheetId !== this.selected.sheetId ||
        previousSelection.startRow !== this.selected.startRow ||
        previousSelection.endRow !== this.selected.endRow ||
        previousSelection.startColumn !== this.selected.startColumn ||
        previousSelection.endColumn !== this.selected.endColumn
      ) {
        this.selectionRevision++;
        this.options.onSelection?.(this.selection);
      }
      if (this.cellEditor.isEditing) this.positionEditor();
      this.schedule();
    });
    this.stopData = book.onData(() => this.schedule());
    this.stopDispose = book.onDispose(() => this.dispose());
    this.resize = new ResizeObserver(() => {
      if (this.cellEditor.isEditing) this.positionEditor();
      this.schedule();
    });
    this.resize.observe(this.element);
    this.schedule();
  }
  get selection(): Selection {
    return { ...this.selected };
  }
  /** Changes whenever the active selection is set or structurally transformed. */
  get selectionVersion() {
    return this.selectionRevision;
  }
  get activeSheetId() {
    return this.sheetId;
  }
  get isComposing() {
    return this.cellEditor.isComposing;
  }
  setReadOnly(value: boolean) {
    this.options.readOnly = value;
    if (value) this.cancelEditing();
    this.element.setAttribute('aria-readonly', String(value));
  }
  setZoom(value: number) {
    if (!Number.isFinite(value))
      throw new OpenSheetError('INVALID_ARGUMENT', 'Zoom must be finite');
    this.zoom = Math.max(0.5, Math.min(2, value));
    this.layoutState.dirty = true;
    if (this.cellEditor.isEditing) this.positionEditor();
    this.schedule();
  }
  getZoom() {
    return this.zoom;
  }
  setSheet(id: string) {
    if (this.cellEditor.isComposing) return;
    if (this.cellEditor.isEditing && !this.finishEditing()) return;
    const sheet = this.book.getSheetById(id);
    if (!sheet) throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown sheet');
    this.sheetId = id;
    this.selectionState.resetMerges();
    this.scroller.scrollTop = 0;
    this.scroller.scrollLeft = 0;
    this.layoutState.dirty = true;
    this.setSelection({ sheetId: id, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
  }
  setSelection(selection: Selection) {
    if (this.disposed) return;
    if (this.applyingDataJump) this.applyingDataJump = false;
    else this.cancelDataJump();
    this.keyboardSelection = undefined;
    if (selection.sheetId !== this.sheetId) {
      this.setSheet(selection.sheetId);
      if (selection.sheetId !== this.sheetId) return;
    }
    this.book.getSheetById(this.sheetId)!.range(selection);
    this.selected = this.expandSelection(selection);
    this.selectionRevision++;
    this.cellEditor.selectionChanged(this.selected);
    if (this.layoutState.dirty) this.layout();
    this.revealCell(this.selected.startRow, this.selected.startColumn);
    this.options.onSelection?.(this.selection);
    this.schedule();
  }
  private mergeAt(row: number, column: number) {
    return this.selectionState.mergeAt(row, column, this.sheetId);
  }
  private expandSelection(selection: Selection) {
    return this.selectionState.expand(selection);
  }
  private selectBetween(a: { row: number; column: number }, b: { row: number; column: number }) {
    this.setSelection({
      sheetId: this.sheetId,
      startRow: Math.min(a.row, b.row),
      endRow: Math.max(a.row, b.row) + 1,
      startColumn: Math.min(a.column, b.column),
      endColumn: Math.max(a.column, b.column) + 1,
    });
  }
  private activeRange() {
    return this.book.getSheetById(this.sheetId)!.range(this.selected);
  }
  focus() {
    this.element.focus();
  }
  private run(fn: () => unknown) {
    try {
      Promise.resolve(fn()).catch((e) => {
        if (!this.disposed) this.options.onError?.(e);
      });
    } catch (e) {
      if (!this.disposed) this.options.onError?.(e);
    }
  }
  private layout() {
    this.layoutState.rebuild(this.book, this.sheetId, this.zoom, this.spacer);
  }
  private position(row: number, column: number) {
    return this.layoutState.position(
      row,
      column,
      this.book.sheetData(this.sheetId).freeze,
      this.scroller,
    );
  }
  private hit(e: MouseEvent) {
    if (this.layoutState.dirty) this.layout();
    return this.layoutState.hit(
      e,
      this.element,
      this.book.sheetData(this.sheetId).freeze,
      this.scroller,
    );
  }
  private schedule() {
    if (this.raf || this.disposed) return;
    this.raf = requestAnimationFrame(() => {
      this.raf = 0;
      this.draw();
    });
  }
  private draw() {
    if (this.disposed) return;
    if (this.layoutState.dirty) this.layout();
    this.scrollbars.update();
    void this.prepareViewport();
    paintGrid(
      this.book,
      this.sheetId,
      this.selected,
      this.zoom,
      this.layoutState,
      this.selectionState,
      this.element,
      this.canvas,
      this.ariaCell,
      this.scroller,
      this.formulaReferences,
    );
  }
  private viewportRanges(): Selection[] {
    const sh = this.book.sheetData(this.sheetId),
      row = this.layoutState.index(this.layoutState.rowOffsets, this.scroller.scrollTop),
      col = this.layoutState.index(this.layoutState.columnOffsets, this.scroller.scrollLeft);
    const range = {
      sheetId: sh.id,
      startRow: Math.max(0, row - 8),
      endRow: Math.min(
        sh.rowOrder.length,
        this.layoutState.index(
          this.layoutState.rowOffsets,
          this.scroller.scrollTop + this.scroller.clientHeight,
        ) + 9,
      ),
      startColumn: Math.max(0, col - 2),
      endColumn: Math.min(
        sh.columnOrder.length,
        this.layoutState.index(
          this.layoutState.columnOffsets,
          this.scroller.scrollLeft + this.scroller.clientWidth,
        ) + 3,
      ),
    };
    const frozenRows = Math.min(
      sh.freeze.rows,
      this.layoutState.index(this.layoutState.rowOffsets, this.scroller.clientHeight) + 1,
    );
    const frozenColumns = Math.min(
      sh.freeze.columns,
      this.layoutState.index(this.layoutState.columnOffsets, this.scroller.clientWidth) + 1,
    );
    const ranges = [range];
    if (sh.freeze.rows) ranges.push({ ...range, startRow: 0, endRow: frozenRows });
    if (sh.freeze.columns)
      ranges.push({
        ...range,
        startColumn: 0,
        endColumn: frozenColumns,
      });
    if (sh.freeze.rows && sh.freeze.columns)
      ranges.push({
        sheetId: sh.id,
        startRow: 0,
        endRow: frozenRows,
        startColumn: 0,
        endColumn: frozenColumns,
      });
    return ranges;
  }
  private prepareViewport(): Promise<void> {
    if (this.fetching) return this.fetching;
    if (this.disposed) return Promise.resolve();
    this.fetching = (async () => {
      try {
        for (const range of this.viewportRanges()) {
          if (this.disposed) break;
          await this.book.prefetch(range.sheetId, range);
        }
      } catch (error) {
        if (!this.disposed) {
          this.fetchError = error;
          this.options.onError?.(error);
        }
      } finally {
        this.fetching = undefined;
      }
    })();
    return this.fetching;
  }
  private async waitForFetch(fetch: Promise<void>) {
    if (this.disposed) return;
    let finish!: () => void;
    const disposed = new Promise<void>((resolve) => {
      finish = resolve;
    });
    this.abort.signal.addEventListener('abort', finish, { once: true });
    try {
      await Promise.race([fetch, disposed]);
    } finally {
      this.abort.signal.removeEventListener('abort', finish);
    }
  }
  /** Optionally record synchronous phases for diagnostics; timings include scheduler delays. */
  async ready(timing?: {
    layoutMs?: number;
    previousFetchMs?: number;
    viewportFetchMs?: number;
    frameWaitMs?: number;
    drawMs?: number;
  }) {
    let phaseStart = timing ? performance.now() : 0;
    if (timing) {
      timing.layoutMs =
        timing.previousFetchMs =
        timing.viewportFetchMs =
        timing.frameWaitMs =
        timing.drawMs =
          0;
    }
    const record = (
      key: 'layoutMs' | 'previousFetchMs' | 'viewportFetchMs' | 'frameWaitMs' | 'drawMs',
    ) => {
      if (!timing) return;
      const now = performance.now();
      timing[key] = timing[key]! + now - phaseStart;
      phaseStart = now;
    };
    while (!this.disposed) {
      if (this.layoutState.dirty) this.layout();
      record('layoutMs');
      while (this.fetching && !this.disposed) await this.waitForFetch(this.fetching);
      if (this.disposed) return;
      if (this.fetchError) {
        const error = this.fetchError;
        this.fetchError = undefined;
        throw error;
      }
      record('previousFetchMs');
      if (this.layoutState.dirty) continue;
      const sheet = this.sheetId,
        revision = this.book.revision,
        zoom = this.zoom,
        top = this.scroller.scrollTop,
        left = this.scroller.scrollLeft,
        width = this.scroller.clientWidth,
        height = this.scroller.clientHeight;
      await this.waitForFetch(this.prepareViewport());
      if (this.disposed) return;
      if (this.fetchError) {
        const error = this.fetchError;
        this.fetchError = undefined;
        throw error;
      }
      record('viewportFetchMs');
      const painted = await new Promise<boolean>((resolve) =>
        requestAnimationFrame(() => {
          record('frameWaitMs');
          if (this.disposed) return resolve(true);
          if (
            this.layoutState.dirty ||
            this.sheetId !== sheet ||
            this.book.revision !== revision ||
            this.zoom !== zoom ||
            this.scroller.scrollTop !== top ||
            this.scroller.scrollLeft !== left ||
            this.scroller.clientWidth !== width ||
            this.scroller.clientHeight !== height
          )
            return resolve(false);
          this.draw();
          record('drawMs');
          resolve(true);
        }),
      );
      if (painted) return;
    }
  }
  startEditing(text?: string) {
    this.cancelDataJump();
    this.cellEditor.start(text);
  }
  private positionEditor() {
    this.cellEditor.position();
    const style = this.cellEditor.element.style;
    Object.assign(this.cellReferenceMirror.style, {
      left: style.left,
      top: style.top,
      width: style.width,
      height: style.height,
    });
    this.syncCellReferenceMirror();
    this.positionAssist();
  }
  private syncCellReferenceMirror() {
    this.cellReferenceMirror.scrollTop = this.cellEditor.element.scrollTop;
    this.cellReferenceMirror.scrollLeft = this.cellEditor.element.scrollLeft;
  }
  private positionAssist() {
    const width = Math.max(0, Math.min(320, this.scroller.clientWidth - 8));
    const left = Math.max(
      0,
      Math.min(this.cellEditor.element.offsetLeft, this.scroller.clientWidth - width),
    );
    const bottom = this.cellEditor.element.offsetTop + this.cellEditor.element.offsetHeight;
    const top = Math.max(0, Math.min(bottom, this.scroller.clientHeight - 150));
    Object.assign(this.formulaAssist.element.style, {
      left: `${left}px`,
      top: `${top}px`,
      right: 'auto',
      width: `${width}px`,
      maxHeight: `${Math.max(0, Math.min(260, this.scroller.clientHeight - top - 4))}px`,
    });
  }
  finishEditing(): boolean {
    return this.cellEditor.finish();
  }
  cancelEditing() {
    this.cellEditor.cancel();
  }
  private cancelDataJump() {
    this.dataJumpEpoch++;
    this.dataJumpQueue = [];
    this.dataJumpAbort?.abort();
    this.dataJumpAbort = undefined;
  }
  private async dataJump(rows: number, columns: number, extend: boolean) {
    if (this.book.persistenceState.pending)
      throw new OpenSheetError('BUSY', 'Wait for saving to finish, then retry the action.');
    if (this.dataJumpQueue.length >= 64)
      throw new OpenSheetError('BUSY', 'Wait for navigation to finish, then retry the action.');
    this.dataJumpQueue.push({ rows, columns, extend });
    const epoch = this.dataJumpEpoch;
    if (this.dataJumpProcessing === epoch) return;
    this.dataJumpProcessing = epoch;
    try {
      while (this.dataJumpEpoch === epoch && this.dataJumpQueue.length) {
        const next = this.dataJumpQueue.shift()!;
        await this.applyDataJump(next.rows, next.columns, next.extend);
      }
    } catch (error) {
      if (this.dataJumpEpoch === epoch) this.cancelDataJump();
      throw error;
    } finally {
      if (this.dataJumpProcessing === epoch) this.dataJumpProcessing = undefined;
    }
  }
  private async applyDataJump(rows: number, columns: number, extend: boolean) {
    if (this.book.persistenceState.pending) {
      this.cancelDataJump();
      return;
    }
    const controller = new AbortController();
    this.dataJumpAbort = controller;
    const selected = this.selected,
      sheet = this.sheetId,
      selectionRevision = this.selectionRevision;
    const focus = extend ? this.keyboardSelection?.focus : undefined;
    const origin = focus ?? {
      row: extend ? selected.endRow - 1 : selected.startRow,
      column: extend ? selected.endColumn - 1 : selected.startColumn,
    };
    try {
      const result = await this.book.navigateData(
        sheet,
        origin.row,
        origin.column,
        rows ? 'row' : 'column',
        (rows || columns) as -1 | 1,
        { signal: controller.signal },
      );
      if (
        this.disposed ||
        controller.signal.aborted ||
        this.dataJumpAbort !== controller ||
        this.cellEditor.isEditing ||
        this.cellEditor.isPending ||
        this.cellEditor.isComposing ||
        this.book.persistenceState.pending ||
        this.sheetId !== sheet ||
        this.selectionRevision !== selectionRevision
      )
        return;
      this.dataJumpAbort = undefined;
      this.applyingDataJump = true;
      try {
        this.move(0, 0, extend, false, result);
      } finally {
        this.applyingDataJump = false;
      }
    } catch (error) {
      if (!controller.signal.aborted) throw error;
    } finally {
      if (this.dataJumpAbort === controller) this.dataJumpAbort = undefined;
    }
  }
  private move(dr: number, dc: number, extend = false, page = false, target?: CellPosition) {
    if (this.layoutState.dirty) this.layout();
    const sh = this.book.getSheetById(this.sheetId)!;
    const old = this.selected;
    const merge = this.mergeAt(old.startRow, old.startColumn);
    const exactMerge =
      merge &&
      old.startRow === merge.startRow &&
      old.endRow === merge.endRow &&
      old.startColumn === merge.startColumn &&
      old.endColumn === merge.endColumn;
    const previousKeyboard = extend ? this.keyboardSelection : undefined;
    const anchor = previousKeyboard?.anchor ?? { row: old.startRow, column: old.startColumn };
    let originRow = extend
      ? (previousKeyboard?.focus.row ?? (exactMerge && dr < 0 ? old.startRow : old.endRow - 1))
      : exactMerge && dr > 0
        ? old.endRow - 1
        : old.startRow;
    let originColumn = extend
      ? (previousKeyboard?.focus.column ??
        (exactMerge && dc < 0 ? old.startColumn : old.endColumn - 1))
      : exactMerge && dc > 0
        ? old.endColumn - 1
        : old.startColumn;
    if (previousKeyboard) {
      const focusedMerge = this.mergeAt(originRow, originColumn);
      if (focusedMerge) {
        if (dr < 0) originRow = focusedMerge.startRow;
        else if (dr > 0) originRow = focusedMerge.endRow - 1;
        if (dc < 0) originColumn = focusedMerge.startColumn;
        else if (dc > 0) originColumn = focusedMerge.endColumn - 1;
      }
    }
    if (target) {
      dr = target.row - originRow;
      dc = target.column - originColumn;
    }
    if (page) {
      const direction = Math.sign(dr);
      const frozen = this.book.sheetData(this.sheetId).freeze.rows;
      const viewport = Math.max(
        1,
        this.scroller.clientHeight - TOP - this.layoutState.rowOffsets[frozen],
      );
      const target = this.layoutState.index(
        this.layoutState.rowOffsets,
        this.layoutState.rowOffsets[originRow] + direction * viewport,
      );
      dr = target === originRow ? direction : target - originRow;
    }
    let row = Math.max(0, Math.min(sh.rowCount - 1, originRow + dr)),
      column = Math.max(0, Math.min(sh.columnCount - 1, originColumn + dc));
    row = this.layoutState.visibleIndex(this.layoutState.rowOffsets, row, dr);
    column = this.layoutState.visibleIndex(this.layoutState.columnOffsets, column, dc);
    const selectionRevision = this.selectionRevision;
    this.selectBetween(extend ? anchor : { row, column }, {
      row,
      column,
    });
    if (this.disposed || this.selectionRevision !== selectionRevision + 1) return;
    if (extend) this.keyboardSelection = { anchor, focus: { row, column } };
    this.revealCell(row, column);
  }
  private revealCell(row: number, column: number) {
    const merge = this.mergeAt(row, column);
    row = merge?.startRow ?? row;
    column = merge?.startColumn ?? column;
    const p = this.position(row, column);
    const frozen = this.book.sheetData(this.sheetId).freeze;
    const frozenHeight = this.layoutState.rowOffsets[frozen.rows];
    const frozenWidth = this.layoutState.columnOffsets[frozen.columns];
    const rowHeight = Math.min(
      this.layoutState.rowOffsets[merge?.endRow ?? row + 1] - this.layoutState.rowOffsets[row],
      Math.max(0, this.scroller.clientHeight - TOP - frozenHeight),
    );
    const columnWidth = Math.min(
      this.layoutState.columnOffsets[merge?.endColumn ?? column + 1] -
        this.layoutState.columnOffsets[column],
      Math.max(0, this.scroller.clientWidth - LEFT - frozenWidth),
    );
    if (merge && row < frozen.rows && merge.endRow > frozen.rows) {
      // The scrollable part starts at the freeze line, so any scroll hides part of it.
      this.scroller.scrollTop = 0;
    } else if (row >= frozen.rows) {
      if (p.y < TOP + frozenHeight)
        this.scroller.scrollTop = Math.max(0, this.layoutState.rowOffsets[row] - frozenHeight);
      else if (p.y + rowHeight > this.scroller.clientHeight)
        this.scroller.scrollTop += p.y + rowHeight - this.scroller.clientHeight;
    }
    if (merge && column < frozen.columns && merge.endColumn > frozen.columns) {
      this.scroller.scrollLeft = 0;
    } else if (column >= frozen.columns) {
      if (p.x < LEFT + frozenWidth)
        this.scroller.scrollLeft = Math.max(
          0,
          this.layoutState.columnOffsets[column] - frozenWidth,
        );
      else if (p.x + columnWidth > this.scroller.clientWidth)
        this.scroller.scrollLeft += p.x + columnWidth - this.scroller.clientWidth;
    }
  }
  moveSelection(rows: number, columns: number, extend = false) {
    if (![rows, columns].every((value) => Number.isInteger(value) && Math.abs(value) <= 1))
      throw new OpenSheetError('INVALID_ARGUMENT', 'Selection movement must be at most one cell');
    this.move(rows, columns, extend);
  }
  takeFormulaDraft() {
    return this.cellEditor.takeFormulaDraft();
  }
  setFormulaReferences(selections: Selection[]) {
    if (selections.length > 64)
      throw new OpenSheetError('LIMIT_EXCEEDED', 'Too many formula references');
    if (!selections.length && !this.formulaReferences.length) return;
    const next = selections.map((selection) => {
      const sheet = this.book.getSheetById(selection.sheetId);
      if (!sheet) throw new OpenSheetError('INVALID_RANGE', 'Unknown sheet');
      sheet.range(selection);
      return { ...selection };
    });
    if (
      next.length === this.formulaReferences.length &&
      next.every((item, i) => {
        const prior = this.formulaReferences[i]!;
        return (
          item.sheetId === prior.sheetId &&
          item.startRow === prior.startRow &&
          item.endRow === prior.endRow &&
          item.startColumn === prior.startColumn &&
          item.endColumn === prior.endColumn
        );
      })
    )
      return;
    this.formulaReferences = next;
    this.schedule();
  }
  copyText() {
    return copyText(this.activeRange());
  }
  copyClipboard() {
    return copyClipboard(this.activeRange());
  }
  pasteText(text: string) {
    if (this.disposed || this.options.readOnly) return;
    const revision = this.selectionRevision;
    this.run(async () => {
      const selection = await pasteText(this.book, this.selected, text, {
        baseRevision: this.book.revision,
      });
      if (!this.disposed && this.selectionRevision === revision) this.setSelection(selection);
    });
  }
  pasteHTML(html: string, fallbackText = '') {
    if (this.disposed || this.options.readOnly) return;
    const revision = this.selectionRevision;
    this.run(async () => {
      const selection = await pasteHTML(this.book, this.selected, html, fallbackText, {
        baseRevision: this.book.revision,
      });
      if (!this.disposed && this.selectionRevision === revision) this.setSelection(selection);
    });
  }
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.cancelDataJump();
    this.cellEditor.dispose();
    this.abort.abort();
    this.input.dispose();
    this.stop();
    this.stopData();
    this.stopDispose();
    this.resize.disconnect();
    cancelAnimationFrame(this.raf);
    this.element.remove();
  }
}
