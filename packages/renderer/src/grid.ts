import { OpenSheetError, Workbook, type Selection } from '@opensheetjs/core';
import { copyText, pasteText } from './clipboard.js';
import { CellEditor } from './editor.js';
import { LEFT, TOP } from './geometry.js';
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
  private fetching?: Promise<void>;
  private fetchError?: unknown;
  private disposed = false;
  private zoom = 1;
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
    );
    this.ariaCell = document.createElement('div');
    this.ariaCell.id = `os-active-${crypto.randomUUID()}`;
    const ariaRow = document.createElement('div');
    ariaRow.className = 'os-sr-only';
    ariaRow.setAttribute('role', 'row');
    ariaRow.append(this.ariaCell);
    this.ariaCell.setAttribute('role', 'gridcell');
    this.element.setAttribute('aria-activedescendant', this.ariaCell.id);
    this.element.append(this.scroller, this.canvas, this.cellEditor.element, ariaRow);
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
      focus: () => this.focus(),
      run: (fn) => this.run(fn),
      range: () => this.activeRange(),
      readOnly: () => !!this.options.readOnly,
      copy: () => this.copyText(),
      paste: (text) => this.pasteText(text),
    });
    this.stop = book.onCommit((commit) => {
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
      if (
        previousSelection.sheetId !== this.selected.sheetId ||
        previousSelection.startRow !== this.selected.startRow ||
        previousSelection.endRow !== this.selected.endRow ||
        previousSelection.startColumn !== this.selected.startColumn ||
        previousSelection.endColumn !== this.selected.endColumn
      )
        this.options.onSelection?.(this.selection);
      if (this.cellEditor.isEditing) this.positionEditor();
      this.schedule();
    });
    this.stopData = book.onData(() => this.schedule());
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
  get activeSheetId() {
    return this.sheetId;
  }
  setReadOnly(value: boolean) {
    this.options.readOnly = value;
    if (value) this.cancelEditing();
    this.element.setAttribute('aria-readonly', String(value));
  }
  setZoom(value: number) {
    this.zoom = Math.max(0.5, Math.min(2, value));
    this.layoutState.dirty = true;
    if (this.cellEditor.isEditing) this.positionEditor();
    this.schedule();
  }
  getZoom() {
    return this.zoom;
  }
  setSheet(id: string) {
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
    if (selection.sheetId !== this.sheetId) {
      this.setSheet(selection.sheetId);
    }
    this.book.getSheetById(this.sheetId)!.range(selection);
    this.selected = this.expandSelection(selection);
    if (this.layoutState.dirty) this.layout();
    const selected = this.selected;
    const p = this.position(selected.startRow, selected.startColumn);
    const f = this.book.sheetData(this.sheetId).freeze;
    if (selected.startRow >= f.rows) {
      if (p.y < TOP + this.layoutState.rowOffsets[f.rows])
        this.scroller.scrollTop = Math.max(
          0,
          this.layoutState.rowOffsets[selected.startRow] - this.layoutState.rowOffsets[f.rows],
        );
      else if (p.y + 30 > this.scroller.clientHeight)
        this.scroller.scrollTop += p.y + 30 - this.scroller.clientHeight;
    }
    if (selected.startColumn >= f.columns) {
      if (p.x < LEFT + this.layoutState.columnOffsets[f.columns])
        this.scroller.scrollLeft = Math.max(
          0,
          this.layoutState.columnOffsets[selected.startColumn] -
            this.layoutState.columnOffsets[f.columns],
        );
      else if (p.x + 80 > this.scroller.clientWidth)
        this.scroller.scrollLeft += p.x + 80 - this.scroller.clientWidth;
    }
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
      Promise.resolve(fn()).catch((e) => this.options.onError?.(e));
    } catch (e) {
      this.options.onError?.(e);
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
        for (const range of this.viewportRanges()) await this.book.prefetch(range.sheetId, range);
      } catch (error) {
        this.fetchError = error;
        this.options.onError?.(error);
      } finally {
        this.fetching = undefined;
      }
    })();
    return this.fetching;
  }
  async ready() {
    if (this.layoutState.dirty) this.layout();
    while (this.fetching) await this.fetching;
    if (this.fetchError) {
      const error = this.fetchError;
      this.fetchError = undefined;
      throw error;
    }
    await this.prepareViewport();
    if (this.fetchError) {
      const error = this.fetchError;
      this.fetchError = undefined;
      throw error;
    }
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => {
        this.draw();
        resolve();
      }),
    );
  }
  startEditing(text?: string) {
    this.cellEditor.start(text);
  }
  private positionEditor() {
    this.cellEditor.position();
  }
  finishEditing(): boolean {
    return this.cellEditor.finish();
  }
  cancelEditing() {
    this.cellEditor.cancel();
  }
  private move(dr: number, dc: number, extend = false) {
    const sh = this.book.getSheetById(this.sheetId)!;
    const old = this.selected;
    const merge = this.mergeAt(old.startRow, old.startColumn);
    const exactMerge =
      merge &&
      old.startRow === merge.startRow &&
      old.endRow === merge.endRow &&
      old.startColumn === merge.startColumn &&
      old.endColumn === merge.endColumn;
    const originRow = extend
      ? old.endRow - 1
      : exactMerge && dr > 0
        ? old.endRow - 1
        : old.startRow;
    const originColumn = extend
      ? old.endColumn - 1
      : exactMerge && dc > 0
        ? old.endColumn - 1
        : old.startColumn;
    let row = Math.max(0, Math.min(sh.rowCount - 1, originRow + dr)),
      column = Math.max(0, Math.min(sh.columnCount - 1, originColumn + dc));
    while (
      dr &&
      row > 0 &&
      row < sh.rowCount - 1 &&
      this.layoutState.rowOffsets[row + 1] === this.layoutState.rowOffsets[row]
    )
      row += Math.sign(dr);
    while (
      dc &&
      column > 0 &&
      column < sh.columnCount - 1 &&
      this.layoutState.columnOffsets[column + 1] === this.layoutState.columnOffsets[column]
    )
      column += Math.sign(dc);
    this.selectBetween(extend ? { row: old.startRow, column: old.startColumn } : { row, column }, {
      row,
      column,
    });
    const p = this.position(row, column);
    if (p.y < TOP) this.scroller.scrollTop = Math.max(0, this.layoutState.rowOffsets[row]);
    else if (p.y + 30 > this.scroller.clientHeight)
      this.scroller.scrollTop += p.y + 30 - this.scroller.clientHeight;
    if (p.x < LEFT) this.scroller.scrollLeft = Math.max(0, this.layoutState.columnOffsets[column]);
    else if (p.x + 128 > this.scroller.clientWidth)
      this.scroller.scrollLeft += p.x + 128 - this.scroller.clientWidth;
  }
  copyText() {
    return copyText(this.activeRange());
  }
  pasteText(text: string) {
    if (this.options.readOnly) return;
    this.run(async () => this.setSelection(await pasteText(this.book, this.selected, text)));
  }
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.abort.abort();
    this.input.dispose();
    this.stop();
    this.stopData();
    this.resize.disconnect();
    cancelAnimationFrame(this.raf);
    this.element.remove();
  }
}
