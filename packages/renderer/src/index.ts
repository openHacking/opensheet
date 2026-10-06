import {
  Workbook,
  address,
  columnName,
  parseInput,
  parseRange,
  rangeAddress,
  contains,
  type Selection,
  type Rect,
  type CellValue,
  OpenSheetError,
} from '@opensheetjs/core';
export type GridOptions = {
  readOnly?: boolean;
  onSelection?: (selection: Selection) => void;
  onError?: (error: unknown) => void;
  onStatus?: (message: string) => void;
};
const LEFT = 48,
  TOP = 30;
export class CanvasGrid {
  readonly element: HTMLDivElement;
  readonly scroller: HTMLDivElement;
  private spacer: HTMLDivElement;
  private canvas: HTMLCanvasElement;
  private editor: HTMLTextAreaElement;
  private ariaCell: HTMLDivElement;
  private sheetId: string;
  private selected: Selection;
  private rowOffsets: number[] = [];
  private columnOffsets: number[] = [];
  private raf = 0;
  private abort = new AbortController();
  private resize: ResizeObserver;
  private stop: () => void;
  private editing = false;
  private composing = false;
  private editPosition = { row: 0, column: 0 };
  private disposed = false;
  private zoom = 1;
  private anchor = { row: 0, column: 0 };
  private drag = false;
  private layoutDirty = true;
  constructor(
    private container: HTMLElement,
    private book: Workbook,
    private options: GridOptions = {},
  ) {
    this.sheetId = book.getSheets().find((s) => !book.sheetData(s.id).hidden)!.id;
    this.selected = { sheetId: this.sheetId, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 };
    this.element = document.createElement('div');
    this.element.className = 'os-grid';
    this.element.tabIndex = 0;
    this.element.setAttribute('role', 'grid');
    this.element.setAttribute('aria-label', 'Spreadsheet');
    this.scroller = document.createElement('div');
    this.scroller.className = 'os-scroll';
    this.spacer = document.createElement('div');
    this.scroller.append(this.spacer);
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'os-canvas';
    this.canvas.setAttribute('aria-hidden', 'true');
    this.editor = document.createElement('textarea');
    this.editor.className = 'os-cell-editor';
    this.editor.setAttribute('aria-label', 'Edit cell');
    this.editor.hidden = true;
    this.ariaCell = document.createElement('div');
    this.ariaCell.id = `os-active-${crypto.randomUUID()}`;
    const ariaRow = document.createElement('div');
    ariaRow.className = 'os-sr-only';
    ariaRow.setAttribute('role', 'row');
    ariaRow.append(this.ariaCell);
    this.ariaCell.setAttribute('role', 'gridcell');
    this.element.setAttribute('aria-activedescendant', this.ariaCell.id);
    this.element.append(this.scroller, this.canvas, this.editor, ariaRow);
    this.element.setAttribute('aria-readonly', String(!!this.options.readOnly));
    container.append(this.element);
    const signal = this.abort.signal;
    this.scroller.addEventListener(
      'scroll',
      () => {
        if (this.editing) this.positionEditor();
        this.schedule();
      },
      { signal },
    );
    this.element.addEventListener('pointerdown', (e) => this.pointer(e), { signal });
    window.addEventListener(
      'pointermove',
      (e) => {
        if (this.drag) {
          const p = this.hit(e);
          if (p) this.selectBetween(this.anchor, p);
        }
      },
      { signal },
    );
    window.addEventListener(
      'pointerup',
      () => {
        this.drag = false;
      },
      { signal },
    );
    this.element.addEventListener(
      'dblclick',
      (e) => {
        if (e.target === this.editor) return;
        this.startEditing();
      },
      { signal },
    );
    this.element.addEventListener('keydown', (e) => this.keydown(e), { signal });
    this.element.addEventListener(
      'compositionstart',
      () => {
        this.composing = true;
        if (!this.editing) this.startEditing('');
      },
      { signal },
    );
    this.element.addEventListener(
      'compositionend',
      () => {
        this.composing = false;
      },
      { signal },
    );
    this.editor.addEventListener(
      'blur',
      () => {
        if (this.editing && !this.composing) this.finishEditing();
      },
      { signal },
    );
    this.element.addEventListener(
      'copy',
      (e) => {
        if (this.editing) return;
        e.clipboardData?.setData('text/plain', this.copyText());
        e.preventDefault();
      },
      { signal },
    );
    this.element.addEventListener(
      'cut',
      (e) => {
        if (this.editing) return;
        if (this.options.readOnly) return;
        e.clipboardData?.setData('text/plain', this.copyText());
        e.preventDefault();
        this.run(() => this.activeRange().clear());
      },
      { signal },
    );
    this.element.addEventListener(
      'paste',
      (e) => {
        if (this.editing) return;
        const text = e.clipboardData?.getData('text/plain');
        if (text !== undefined) {
          e.preventDefault();
          this.pasteText(text);
        }
      },
      { signal },
    );
    this.stop = book.onCommit(() => {
      if (!book.getSheetById(this.sheetId))
        this.setSheet(book.getSheets().find((s) => !book.sheetData(s.id).hidden)!.id);
      this.layoutDirty = true;
      const sh = book.getSheetById(this.sheetId)!;
      this.selected = {
        ...this.selected,
        startRow: Math.min(this.selected.startRow, sh.rowCount - 1),
        endRow: Math.min(this.selected.endRow, sh.rowCount),
        startColumn: Math.min(this.selected.startColumn, sh.columnCount - 1),
        endColumn: Math.min(this.selected.endColumn, sh.columnCount),
      };
      this.schedule();
    });
    this.resize = new ResizeObserver(() => this.schedule());
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
    this.layoutDirty = true;
    this.schedule();
  }
  getZoom() {
    return this.zoom;
  }
  setSheet(id: string) {
    if (this.editing && !this.finishEditing()) return;
    const sheet = this.book.getSheetById(id);
    if (!sheet) throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown sheet');
    this.sheetId = id;
    this.scroller.scrollTop = 0;
    this.scroller.scrollLeft = 0;
    this.layoutDirty = true;
    this.setSelection({ sheetId: id, startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 });
  }
  setSelection(selection: Selection) {
    if (selection.sheetId !== this.sheetId) {
      this.setSheet(selection.sheetId);
    }
    this.book.getSheetById(this.sheetId)!.range(selection);
    this.selected = { ...selection };
    if (this.layoutDirty) this.layout();
    const p = this.position(selection.startRow, selection.startColumn);
    const f = this.book.sheetData(this.sheetId).freeze;
    if (selection.startRow >= f.rows) {
      if (p.y < TOP + this.rowOffsets[f.rows])
        this.scroller.scrollTop = Math.max(
          0,
          this.rowOffsets[selection.startRow] - this.rowOffsets[f.rows],
        );
      else if (p.y + 30 > this.element.clientHeight)
        this.scroller.scrollTop += p.y + 30 - this.element.clientHeight;
    }
    if (selection.startColumn >= f.columns) {
      if (p.x < LEFT + this.columnOffsets[f.columns])
        this.scroller.scrollLeft = Math.max(
          0,
          this.columnOffsets[selection.startColumn] - this.columnOffsets[f.columns],
        );
      else if (p.x + 80 > this.element.clientWidth)
        this.scroller.scrollLeft += p.x + 80 - this.element.clientWidth;
    }
    this.options.onSelection?.(this.selection);
    this.schedule();
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
      fn();
    } catch (e) {
      this.options.onError?.(e);
    }
  }
  private layout() {
    const sh = this.book.sheetData(this.sheetId);
    this.rowOffsets = [0];
    this.columnOffsets = [0];
    for (let i = 0; i < sh.rowOrder.length; i++) {
      const m = sh.rows[sh.rowOrder[i]];
      const filtered =
        sh.filter &&
        i > 0 &&
        !this.book
          .display(this.sheetId, i, sh.filter.column)
          .toLowerCase()
          .includes(sh.filter.query.toLowerCase());
      this.rowOffsets.push(
        this.rowOffsets[i] + (m?.hidden || filtered ? 0 : (m?.size ?? 30) * this.zoom),
      );
    }
    for (let i = 0; i < sh.columnOrder.length; i++) {
      const m = sh.columns[sh.columnOrder[i]];
      this.columnOffsets.push(
        this.columnOffsets[i] + (m?.hidden ? 0 : (m?.size ?? 128) * this.zoom),
      );
    }
    this.spacer.style.width = `${LEFT + this.columnOffsets.at(-1)!}px`;
    this.spacer.style.height = `${TOP + this.rowOffsets.at(-1)!}px`;
    this.layoutDirty = false;
  }
  private position(row: number, column: number) {
    const freeze = this.book.sheetData(this.sheetId).freeze;
    return {
      x:
        LEFT +
        this.columnOffsets[column] -
        (column >= freeze.columns ? this.scroller.scrollLeft : 0),
      y: TOP + this.rowOffsets[row] - (row >= freeze.rows ? this.scroller.scrollTop : 0),
    };
  }
  private index(offsets: number[], value: number) {
    let lo = 0,
      hi = offsets.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (offsets[mid] <= value) lo = mid;
      else hi = mid - 1;
    }
    return Math.min(lo, offsets.length - 2);
  }
  private hit(e: PointerEvent) {
    if (this.layoutDirty) this.layout();
    const rect = this.element.getBoundingClientRect(),
      x = e.clientX - rect.left - LEFT,
      y = e.clientY - rect.top - TOP;
    if (x < 0 || y < 0 || e.clientX > rect.right || e.clientY > rect.bottom) return null;
    const f = this.book.sheetData(this.sheetId).freeze;
    return {
      row: this.index(
        this.rowOffsets,
        y + (y >= this.rowOffsets[f.rows] ? this.scroller.scrollTop : 0),
      ),
      column: this.index(
        this.columnOffsets,
        x + (x >= this.columnOffsets[f.columns] ? this.scroller.scrollLeft : 0),
      ),
    };
  }
  private pointer(e: PointerEvent) {
    if (e.target === this.editor) return;
    if (this.editing && !this.finishEditing()) return;
    const p = this.hit(e);
    if (!p) return;
    this.focus();
    this.anchor = e.shiftKey
      ? { row: this.selected.startRow, column: this.selected.startColumn }
      : p;
    this.drag = true;
    this.selectBetween(this.anchor, p);
    e.preventDefault();
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
    if (this.layoutDirty) this.layout();
    const w = this.element.clientWidth,
      h = this.element.clientHeight,
      dpr = window.devicePixelRatio || 1;
    if (!w || !h) return;
    this.canvas.width = Math.round(w * dpr);
    this.canvas.height = Math.round(h * dpr);
    this.canvas.style.width = `${w}px`;
    this.canvas.style.height = `${h}px`;
    const ctx = this.canvas.getContext('2d')!;
    ctx.scale(dpr, dpr);
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, w, h);
    const sh = this.book.sheetData(this.sheetId),
      f = sh.freeze;
    const visible = (offsets: number[], scroll: number, extent: number, frozen: number) => {
      const out: number[] = [];
      for (let i = 0; i < Math.min(frozen, offsets.length - 1) && offsets[i] < extent; i++)
        if (offsets[i + 1] > offsets[i]) out.push(i);
      const start = Math.max(frozen, this.index(offsets, scroll + offsets[frozen]));
      for (let i = start; i < offsets.length - 1 && offsets[i] - scroll < extent; i++)
        if (offsets[i + 1] > offsets[i]) out.push(i);
      return out;
    };
    const rows = visible(this.rowOffsets, this.scroller.scrollTop, h - TOP, f.rows),
      cols = visible(this.columnOffsets, this.scroller.scrollLeft, w - LEFT, f.columns);
    const drawCell = (r: number, c: number) => {
      const merge = sh.merges.find(
        (m) => r >= m.startRow && r < m.endRow && c >= m.startColumn && c < m.endColumn,
      );
      if (merge && (r !== merge.startRow || c !== merge.startColumn)) return;
      const pos = this.position(r, c),
        cw = this.columnOffsets[merge?.endColumn ?? c + 1] - this.columnOffsets[c],
        rh = this.rowOffsets[merge?.endRow ?? r + 1] - this.rowOffsets[r];
      const cell = this.book.getCell(this.sheetId, r, c),
        style = this.book.getStyle(cell?.styleId);
      ctx.save();
      const clipX = c >= f.columns ? LEFT + this.columnOffsets[f.columns] : LEFT,
        clipY = r >= f.rows ? TOP + this.rowOffsets[f.rows] : TOP;
      ctx.beginPath();
      ctx.rect(clipX, clipY, w - clipX, h - clipY);
      ctx.clip();
      ctx.fillStyle = style.background ?? '#ffffff';
      ctx.fillRect(pos.x, pos.y, cw, rh);
      if (
        r >= this.selected.startRow &&
        r < this.selected.endRow &&
        c >= this.selected.startColumn &&
        c < this.selected.endColumn
      ) {
        ctx.fillStyle = 'rgba(16,130,94,.06)';
        ctx.fillRect(pos.x, pos.y, cw, rh);
      }
      ctx.strokeStyle = style.border ? '#8ca098' : '#e9eeeb';
      ctx.lineWidth = 1;
      ctx.strokeRect(pos.x + 0.5, pos.y + 0.5, cw, rh);
      ctx.beginPath();
      ctx.rect(pos.x + 3, pos.y + 1, Math.max(0, cw - 6), Math.max(0, rh - 2));
      ctx.clip();
      ctx.fillStyle = style.color ?? '#293d35';
      ctx.font = `${style.italic ? 'italic ' : ''}${style.bold ? '600 ' : ''}${(style.fontSize ?? 13) * this.zoom}px Inter, -apple-system, sans-serif`;
      ctx.textBaseline = 'middle';
      const raw = this.book.value(this.sheetId, r, c);
      const align = style.align ?? (typeof raw === 'number' ? 'right' : 'left');
      ctx.textAlign = align;
      const x =
        align === 'right' ? pos.x + cw - 10 : align === 'center' ? pos.x + cw / 2 : pos.x + 10;
      const text = this.book.display(this.sheetId, r, c);
      if (style.wrap) {
        const lineHeight = (style.fontSize ?? 13) * this.zoom * 1.35;
        const lines: string[] = [];
        let line = '';
        for (const char of text.slice(0, 4096)) {
          if (char === '\n' || ctx.measureText(line + char).width > cw - 20) {
            lines.push(line);
            line = char === '\n' ? '' : char;
          } else line += char;
          if (lines.length > Math.ceil(rh / lineHeight)) break;
        }
        if (line) lines.push(line);
        lines.forEach((part, i) => ctx.fillText(part, x, pos.y + lineHeight * (i + 0.6)));
      } else ctx.fillText(text.replace(/\r?\n/g, ' '), x, pos.y + rh / 2);
      if (style.underline) {
        const length = Math.min(ctx.measureText(text).width, cw - 20);
        ctx.fillRect(
          align === 'right' ? x - length : align === 'center' ? x - length / 2 : x,
          pos.y + rh / 2 + 8,
          length,
          1,
        );
      }
      ctx.restore();
    };
    // Include offscreen merge anchors whose merged rectangle intersects the viewport.
    const anchors = new Set<string>();
    for (const r of rows)
      for (const c of cols) {
        const m = sh.merges.find(
          (m) => r >= m.startRow && r < m.endRow && c >= m.startColumn && c < m.endColumn,
        );
        if (m) anchors.add(`${m.startRow},${m.startColumn}`);
        else drawCell(r, c);
      }
    for (const a of anchors) {
      const [r, c] = a.split(',').map(Number);
      drawCell(r, c);
    }
    ctx.save();
    ctx.beginPath();
    ctx.rect(LEFT, TOP, w - LEFT, h - TOP);
    ctx.clip();
    const s = this.selected,
      start = this.position(s.startRow, s.startColumn),
      end = this.position(s.endRow - 1, s.endColumn - 1);
    ctx.strokeStyle = '#168462';
    ctx.lineWidth = 2;
    ctx.strokeRect(
      start.x + 1,
      start.y + 1,
      end.x + this.columnOffsets[s.endColumn] - this.columnOffsets[s.endColumn - 1] - start.x - 2,
      end.y + this.rowOffsets[s.endRow] - this.rowOffsets[s.endRow - 1] - start.y - 2,
    );
    ctx.restore();
    ctx.fillStyle = '#f6f8f6';
    ctx.fillRect(0, 0, w, TOP);
    ctx.fillRect(0, 0, LEFT, h);
    ctx.font = '11px Inter, sans-serif';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.strokeStyle = '#e3e9e5';
    for (const c of cols) {
      const { x } = this.position(0, c),
        cw = this.columnOffsets[c + 1] - this.columnOffsets[c];
      if (x < LEFT) continue;
      ctx.fillStyle = c >= s.startColumn && c < s.endColumn ? '#deeee6' : '#f6f8f6';
      ctx.fillRect(x, 0, cw, TOP);
      ctx.fillStyle = '#667b6f';
      ctx.fillText(columnName(c), x + cw / 2, TOP / 2);
      ctx.strokeRect(x + 0.5, 0.5, cw, TOP);
    }
    for (const r of rows) {
      const { y } = this.position(r, 0),
        rh = this.rowOffsets[r + 1] - this.rowOffsets[r];
      if (y < TOP) continue;
      ctx.fillStyle = r >= s.startRow && r < s.endRow ? '#deeee6' : '#f6f8f6';
      ctx.fillRect(0, y, LEFT, rh);
      ctx.fillStyle = '#667b6f';
      ctx.fillText(String(r + 1), LEFT / 2, y + rh / 2);
      ctx.strokeRect(0.5, y + 0.5, LEFT, rh);
    }
    if (f.rows) {
      ctx.strokeStyle = '#a7bdb0';
      ctx.beginPath();
      ctx.moveTo(LEFT, TOP + this.rowOffsets[f.rows]);
      ctx.lineTo(w, TOP + this.rowOffsets[f.rows]);
      ctx.stroke();
    }
    if (f.columns) {
      ctx.strokeStyle = '#a7bdb0';
      ctx.beginPath();
      ctx.moveTo(LEFT + this.columnOffsets[f.columns], TOP);
      ctx.lineTo(LEFT + this.columnOffsets[f.columns], h);
      ctx.stroke();
    }
    this.element.setAttribute('aria-rowcount', String(sh.rowOrder.length));
    this.element.setAttribute('aria-colcount', String(sh.columnOrder.length));
    this.ariaCell.setAttribute('aria-rowindex', String(s.startRow + 1));
    this.ariaCell.setAttribute('aria-colindex', String(s.startColumn + 1));
    this.ariaCell.textContent = `${address(s.startRow, s.startColumn)}: ${this.book.display(this.sheetId, s.startRow, s.startColumn)}`;
  }
  startEditing(text?: string) {
    if (this.options.readOnly) return;
    const s = this.selected,
      sh = this.book.sheetData(this.sheetId),
      m = sh.merges.find(
        (m) =>
          s.startRow >= m.startRow &&
          s.startRow < m.endRow &&
          s.startColumn >= m.startColumn &&
          s.startColumn < m.endColumn,
      );
    this.editPosition = { row: m?.startRow ?? s.startRow, column: m?.startColumn ?? s.startColumn };
    const c = this.book.getCell(this.sheetId, this.editPosition.row, this.editPosition.column);
    this.editor.value =
      text ??
      (c?.input.type === 'formula'
        ? '=' + c.input.expression
        : c?.input.type === 'error'
          ? c.input.code
          : c?.input.type === 'blank'
            ? ''
            : String(c && 'value' in c.input ? c.input.value : ''));
    this.editing = true;
    this.editor.hidden = false;
    this.positionEditor();
    this.editor.focus();
    if (text === undefined) this.editor.select();
  }
  private positionEditor() {
    if (this.layoutDirty) this.layout();
    const { row, column } = this.editPosition,
      p = this.position(row, column);
    Object.assign(this.editor.style, {
      left: `${p.x}px`,
      top: `${p.y}px`,
      width: `${Math.max(100, this.columnOffsets[column + 1] - this.columnOffsets[column])}px`,
      height: `${Math.max(30, this.rowOffsets[row + 1] - this.rowOffsets[row])}px`,
    });
  }
  finishEditing(): boolean {
    if (!this.editing) return true;
    try {
      this.book
        .getSheetById(this.sheetId)!
        .range(address(this.editPosition.row, this.editPosition.column))
        .setInput(this.editor.value);
      this.cancelEditing();
      return true;
    } catch (e) {
      this.options.onError?.(e);
      this.editor.focus();
      return false;
    }
  }
  cancelEditing() {
    this.editing = false;
    this.editor.hidden = true;
  }
  private move(dr: number, dc: number, extend = false) {
    const sh = this.book.getSheetById(this.sheetId)!;
    const old = this.selected;
    let row = Math.max(0, Math.min(sh.rowCount - 1, (extend ? old.endRow - 1 : old.startRow) + dr)),
      column = Math.max(
        0,
        Math.min(sh.columnCount - 1, (extend ? old.endColumn - 1 : old.startColumn) + dc),
      );
    while (
      dr &&
      row > 0 &&
      row < sh.rowCount - 1 &&
      this.rowOffsets[row + 1] === this.rowOffsets[row]
    )
      row += Math.sign(dr);
    while (
      dc &&
      column > 0 &&
      column < sh.columnCount - 1 &&
      this.columnOffsets[column + 1] === this.columnOffsets[column]
    )
      column += Math.sign(dc);
    this.selectBetween(extend ? { row: old.startRow, column: old.startColumn } : { row, column }, {
      row,
      column,
    });
    const p = this.position(row, column);
    if (p.y < TOP) this.scroller.scrollTop = Math.max(0, this.rowOffsets[row]);
    else if (p.y + 30 > this.element.clientHeight)
      this.scroller.scrollTop += p.y + 30 - this.element.clientHeight;
    if (p.x < LEFT) this.scroller.scrollLeft = Math.max(0, this.columnOffsets[column]);
    else if (p.x + 128 > this.element.clientWidth)
      this.scroller.scrollLeft += p.x + 128 - this.element.clientWidth;
  }
  private keydown(e: KeyboardEvent) {
    if (e.isComposing || this.composing || e.keyCode === 229) return;
    if (this.editing) {
      if (e.key === 'Escape') {
        e.preventDefault();
        this.cancelEditing();
        this.focus();
      } else if ((e.key === 'Enter' && !e.altKey) || e.key === 'Tab') {
        e.preventDefault();
        if (this.finishEditing()) {
          this.focus();
          this.move(
            e.key === 'Enter' ? (e.shiftKey ? -1 : 1) : 0,
            e.key === 'Tab' ? (e.shiftKey ? -1 : 1) : 0,
          );
        }
      }
      return;
    }
    const mod = e.ctrlKey || e.metaKey;
    if (mod && e.key.toLowerCase() === 'd') {
      e.preventDefault();
      this.run(() => this.activeRange().fillDown());
      return;
    }
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      this.run(() => (e.shiftKey ? this.book.redo() : this.book.undo()));
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      this.run(() => this.book.redo());
      return;
    }
    if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      this.setSelection({
        sheetId: this.sheetId,
        ...this.book.getSheetById(this.sheetId)!.getUsedRange().bounds,
      });
      return;
    }
    const moves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
      Tab: [0, e.shiftKey ? -1 : 1],
    };
    if (moves[e.key]) {
      e.preventDefault();
      this.move(...moves[e.key], e.key !== 'Tab' && e.shiftKey);
      return;
    }
    if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      this.startEditing();
      return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      if (!this.options.readOnly) this.run(() => this.activeRange().clear());
      return;
    }
    if (e.key === 'Home') {
      e.preventDefault();
      this.setSelection({
        sheetId: this.sheetId,
        startRow: mod ? 0 : this.selected.startRow,
        endRow: mod ? 1 : this.selected.startRow + 1,
        startColumn: 0,
        endColumn: 1,
      });
      return;
    }
    if (e.key.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      this.startEditing(e.key);
    }
  }
  copyText() {
    return this.activeRange()
      .getDisplayValues()
      .map((row) =>
        row.map((v) => (/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join('\t'),
      )
      .join('\n');
  }
  pasteText(text: string) {
    if (this.options.readOnly) return;
    this.run(() => {
      if (text.length > 2 * 1024 * 1024)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste is too large');
      const rows: string[][] = [];
      let row: string[] = [],
        cell = '',
        quoted = false;
      for (let i = 0; i < text.length; i++) {
        const c = text[i];
        if (c === '"' && (quoted || !cell)) {
          if (quoted && text[i + 1] === '"') {
            cell += '"';
            i++;
          } else quoted = !quoted;
        } else if (!quoted && (c === '\t' || c === '\n' || c === '\r')) {
          row.push(cell);
          cell = '';
          if (c !== '\t') {
            rows.push(row);
            row = [];
            if (c === '\r' && text[i + 1] === '\n') i++;
          }
        } else cell += c;
      }
      if (quoted) throw new OpenSheetError('INVALID_ARGUMENT', 'Unclosed quote');
      if (cell || row.length || !rows.length) {
        row.push(cell);
        rows.push(row);
      }
      const width = Math.max(...rows.map((r) => r.length));
      const cells = rows.flatMap((row, i) =>
        Array.from({ length: width }, (_, j) => ({
          row: this.selected.startRow + i,
          column: this.selected.startColumn + j,
          input: parseInput(row[j] ?? ''),
        })),
      );
      this.book.execute({ type: 'core.cells.set', payload: { sheetId: this.sheetId, cells } });
      this.setSelection({
        ...this.selected,
        endRow: this.selected.startRow + rows.length,
        endColumn: this.selected.startColumn + width,
      });
    });
  }
  dispose() {
    if (this.disposed) return;
    this.disposed = true;
    this.abort.abort();
    this.stop();
    this.resize.disconnect();
    cancelAnimationFrame(this.raf);
    this.element.remove();
  }
}
