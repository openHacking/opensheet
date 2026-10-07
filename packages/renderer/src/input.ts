import type { Range, Selection, Workbook } from '@opensheetjs/core';
import type { CellEditor } from './editor.js';
import type { CellPosition } from './geometry.js';
export interface InputActions {
  selection(): Selection;
  select(selection: Selection): void;
  selectBetween(a: CellPosition, b: CellPosition): void;
  hit(event: MouseEvent): CellPosition | null;
  move(rows: number, columns: number, extend?: boolean): void;
  focus(): void;
  run(action: () => unknown): void;
  range(): Range;
  readOnly(): boolean;
  copy(): Promise<string>;
  paste(text: string): void;
}
export class GridInput {
  private abort = new AbortController();
  private anchor = { row: 0, column: 0 };
  private drag = false;
  constructor(
    element: HTMLElement,
    private book: Workbook,
    private editor: CellEditor,
    private actions: InputActions,
  ) {
    const signal = this.abort.signal;
    element.addEventListener('pointerdown', (e) => this.pointer(e), { signal });
    window.addEventListener(
      'pointermove',
      (e) => {
        if (this.drag) {
          const p = actions.hit(e);
          if (p) actions.selectBetween(this.anchor, p);
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
    element.addEventListener(
      'dblclick',
      (e) => {
        if (e.target !== editor.element && actions.hit(e)) editor.start();
      },
      { signal },
    );
    element.addEventListener('keydown', (e) => this.keydown(e), { signal });
    element.addEventListener(
      'compositionstart',
      () => {
        editor.isComposing = true;
        if (!editor.isEditing) editor.start('');
      },
      { signal },
    );
    element.addEventListener(
      'compositionend',
      () => {
        editor.isComposing = false;
      },
      { signal },
    );
    editor.element.addEventListener(
      'blur',
      () => {
        if (editor.isEditing && !editor.isComposing) editor.finish();
      },
      { signal },
    );
    editor.element.addEventListener('input', () => editor.position(), { signal });
    element.addEventListener(
      'copy',
      (e) => {
        if (editor.isEditing) return;
        actions.run(async () => {
          const text = await actions.copy();
          await navigator.clipboard.writeText(text);
        });
        e.preventDefault();
      },
      { signal },
    );
    element.addEventListener(
      'cut',
      (e) => {
        if (editor.isEditing || actions.readOnly()) return;
        actions.run(async () => {
          const text = await actions.copy();
          await navigator.clipboard.writeText(text);
        });
        e.preventDefault();
        actions.run(() => actions.range().clear());
      },
      { signal },
    );
    element.addEventListener(
      'paste',
      (e) => {
        if (editor.isEditing) return;
        const text = e.clipboardData?.getData('text/plain');
        if (text !== undefined) {
          e.preventDefault();
          actions.paste(text);
        }
      },
      { signal },
    );
  }
  dispose() {
    this.abort.abort();
  }
  pointer(e: PointerEvent) {
    if (e.target === this.editor.element) return;
    if (this.editor.isEditing && !this.editor.finish()) return;
    const p = this.actions.hit(e);
    if (!p) return;
    this.actions.focus();
    this.anchor = e.shiftKey
      ? { row: this.actions.selection().startRow, column: this.actions.selection().startColumn }
      : p;
    this.drag = true;
    this.actions.selectBetween(this.anchor, p);
    e.preventDefault();
  }
  keydown(e: KeyboardEvent) {
    if (e.isComposing || this.editor.isComposing || e.keyCode === 229) return;
    if (this.editor.isEditing) {
      if (e.key === 'Escape') {
        e.preventDefault();
        this.editor.cancel();
        this.actions.focus();
      } else if ((e.key === 'Enter' && !e.altKey) || e.key === 'Tab') {
        e.preventDefault();
        if (this.editor.finish()) {
          this.actions.focus();
          this.actions.move(
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
      this.actions.run(() => this.actions.range().fillDown());
      return;
    }
    if (mod && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      this.actions.run(() => (e.shiftKey ? this.book.redo() : this.book.undo()));
      return;
    }
    if (mod && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      this.actions.run(() => this.book.redo());
      return;
    }
    if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      this.actions.select({
        sheetId: this.actions.selection().sheetId,
        ...this.book.getSheetById(this.actions.selection().sheetId)!.getUsedRange().bounds,
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
      this.actions.move(...moves[e.key], e.key !== 'Tab' && e.shiftKey);
      return;
    }
    if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      this.editor.start();
      return;
    }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      if (!this.actions.readOnly()) this.actions.run(() => this.actions.range().clear());
      return;
    }
    if (e.key === 'Home') {
      e.preventDefault();
      this.actions.select({
        sheetId: this.actions.selection().sheetId,
        startRow: mod ? 0 : this.actions.selection().startRow,
        endRow: mod ? 1 : this.actions.selection().startRow + 1,
        startColumn: 0,
        endColumn: 1,
      });
      return;
    }
    if (e.key.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      this.editor.start(e.key);
    }
  }
}
