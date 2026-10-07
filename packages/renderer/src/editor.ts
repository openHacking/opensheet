import { address, type Selection, type Workbook } from '@opensheetjs/core';
import type { GridLayout } from './layout.js';
import type { GridSelection } from './selection.js';
export class CellEditor {
  readonly element = document.createElement('textarea');
  isEditing = false;
  isComposing = false;
  private editPosition = { row: 0, column: 0 };
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
  ) {
    this.element.className = 'os-cell-editor';
    this.element.setAttribute('aria-label', 'Edit cell');
    this.element.hidden = true;
  }
  start(text?: string) {
    if (this.readOnly()) return;
    const s = this.getSelection(),
      m = this.selection.mergeAt(s.startRow, s.startColumn);
    this.editPosition = { row: m?.startRow ?? s.startRow, column: m?.startColumn ?? s.startColumn };
    const c = this.book.getCell(this.getSheetId(), this.editPosition.row, this.editPosition.column);
    this.element.value =
      text ??
      (c?.input.type === 'formula'
        ? '=' + c.input.expression
        : c?.input.type === 'error'
          ? c.input.code
          : c?.input.type === 'blank'
            ? ''
            : String(c && 'value' in c.input ? c.input.value : ''));
    this.isEditing = true;
    this.element.hidden = false;
    this.position();
    this.element.focus();
    if (text === undefined) this.element.select();
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
      this.book
        .getSheetById(this.getSheetId())!
        .range(address(this.editPosition.row, this.editPosition.column))
        .setInput(this.element.value);
      this.cancel();
      return true;
    } catch (e) {
      this.onError(e);
      this.element.focus();
      return false;
    }
  }
  cancel() {
    this.isEditing = false;
    this.element.hidden = true;
  }
}
