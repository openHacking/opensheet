import { OpenSheetError, type Range, type Selection, type Workbook } from '@opensheetjs/core';
import type { CellEditor } from './editor.js';
import type { CellPosition } from './geometry.js';
import { cutText } from './clipboard.js';
import { writeClipboard } from './copy-clipboard.js';
export interface InputActions {
  selection(): Selection;
  select(selection: Selection): void;
  selectBetween(a: CellPosition, b: CellPosition): void;
  hit(event: MouseEvent): CellPosition | null;
  selectHeader(event: PointerEvent): boolean;
  move(rows: number, columns: number, extend?: boolean): void;
  page(direction: -1 | 1, extend: boolean): void;
  dataJump(rows: number, columns: number, extend: boolean): void;
  cancelDataJump(): void;
  end(extend: boolean): void;
  focus(): void;
  run(action: () => unknown): void;
  range(): Range;
  readOnly(): boolean;
  paste(text: string, html?: string): void;
  positionEditor(): void;
  startKeyboardReference(): boolean;
  referencePicker?: {
    begin(): boolean;
    update(selection: Selection): void;
    end(): void;
  };
}
export class GridInput {
  private abort = new AbortController();
  private anchor = { row: 0, column: 0 };
  private drag = false;
  private referenceDrag: 'bar' | 'cell' | undefined;
  private compositionBlurred = false;
  private deferredCompositionFinish?: ReturnType<typeof setTimeout>;
  constructor(
    element: HTMLElement,
    private book: Workbook,
    private editor: CellEditor,
    private actions: InputActions,
  ) {
    const signal = this.abort.signal;
    element.addEventListener('pointerdown', (e) => this.pointer(e), { signal });
    element.addEventListener(
      'focusout',
      (event) => {
        if (!(event.relatedTarget instanceof Node && element.contains(event.relatedTarget)))
          actions.cancelDataJump();
        if (
          !this.editor.isEditing &&
          !(event.relatedTarget instanceof Node && element.contains(event.relatedTarget))
        )
          this.editor.cancelPending();
      },
      { signal },
    );
    window.addEventListener(
      'pointermove',
      (e) => {
        if (this.drag) {
          const p = actions.hit(e);
          if (p) {
            actions.selectBetween(this.anchor, p);
            if (this.referenceDrag === 'bar') actions.referencePicker?.update(actions.selection());
            if (this.referenceDrag === 'cell') editor.updateReference(actions.selection());
          }
        }
      },
      { signal },
    );
    window.addEventListener(
      'pointerup',
      () => {
        this.drag = false;
        if (this.referenceDrag === 'bar') actions.referencePicker?.end();
        if (this.referenceDrag === 'cell') editor.endReference();
        this.referenceDrag = undefined;
      },
      { signal },
    );
    window.addEventListener(
      'pointercancel',
      () => {
        this.drag = false;
        if (this.referenceDrag === 'bar') actions.referencePicker?.end();
        if (this.referenceDrag === 'cell') editor.endReference();
        this.referenceDrag = undefined;
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
        actions.cancelDataJump();
        this.clearDeferredCompositionFinish();
        this.compositionBlurred = false;
        editor.isComposing = true;
        if (!editor.isEditing) editor.start('');
      },
      { signal },
    );
    element.addEventListener(
      'compositionend',
      (event) => {
        editor.isComposing = false;
        if (event.target === editor.element && this.compositionBlurred) {
          this.compositionBlurred = false;
          // Some engines deliver the final input after compositionend. Let that event
          // update the draft before committing a blurred editor.
          this.deferredCompositionFinish = setTimeout(() => {
            this.deferredCompositionFinish = undefined;
            if (
              editor.isEditing &&
              !editor.isComposing &&
              document.activeElement !== editor.element
            )
              editor.finish();
          }, 0);
        }
      },
      { signal },
    );
    editor.element.addEventListener(
      'blur',
      () => {
        if (editor.isEditing && editor.isComposing) this.compositionBlurred = true;
        else if (editor.isEditing) editor.finish();
      },
      { signal },
    );
    editor.element.addEventListener(
      'focus',
      () => {
        this.clearDeferredCompositionFinish();
        if (this.compositionBlurred) {
          this.compositionBlurred = false;
          editor.isComposing = false;
        }
      },
      { signal },
    );
    editor.element.addEventListener('input', () => actions.positionEditor(), { signal });
    element.addEventListener(
      'copy',
      (e) => {
        if (editor.isEditing) return;
        const range = actions.range();
        actions.run(() => writeClipboard(range));
        e.preventDefault();
      },
      { signal },
    );
    element.addEventListener(
      'cut',
      (e) => {
        if (editor.isEditing || actions.readOnly()) return;
        const range = actions.range();
        e.preventDefault();
        actions.run(async () => {
          try {
            await cutText(
              book,
              range,
              (text) => navigator.clipboard.writeText(text),
              () => !signal.aborted && !actions.readOnly(),
            );
          } catch (error) {
            if (!signal.aborted) throw error;
          }
        });
      },
      { signal },
    );
    element.addEventListener(
      'paste',
      (e) => {
        if (editor.isEditing) return;
        if (!e.clipboardData?.types.some((type) => type === 'text/plain' || type === 'text/html'))
          return;
        const text = e.clipboardData?.getData('text/plain');
        const html = e.clipboardData?.getData('text/html');
        if (text !== undefined) {
          e.preventDefault();
          actions.paste(text, html);
        }
      },
      { signal },
    );
  }
  dispose() {
    this.clearDeferredCompositionFinish();
    this.abort.abort();
  }
  private clearDeferredCompositionFinish() {
    if (this.deferredCompositionFinish !== undefined) clearTimeout(this.deferredCompositionFinish);
    this.deferredCompositionFinish = undefined;
  }
  pointer(e: PointerEvent) {
    if (e.button !== 0 || e.ctrlKey || (e.target as Element).closest('.os-formula-assist')) return;
    if (e.target === this.editor.element) return;
    this.editor.cancelPending();
    if (this.editor.isComposing) {
      e.preventDefault();
      return;
    }
    const p = this.actions.hit(e);
    if (p && e.button === 0 && this.editor.beginReference()) {
      this.anchor = p;
      this.drag = true;
      this.referenceDrag = 'cell';
      this.actions.selectBetween(p, p);
      this.editor.updateReference(this.actions.selection());
      e.preventDefault();
      return;
    }
    if (this.editor.isEditing && !this.editor.finish()) return;
    if (!p) {
      if (this.actions.selectHeader(e)) e.preventDefault();
      return;
    }
    if (e.button === 0 && this.actions.referencePicker?.begin()) {
      this.anchor = p;
      this.drag = true;
      this.referenceDrag = 'bar';
      this.actions.selectBetween(p, p);
      this.actions.referencePicker.update(this.actions.selection());
      e.preventDefault();
      return;
    }
    this.actions.focus();
    this.anchor = e.shiftKey
      ? { row: this.actions.selection().startRow, column: this.actions.selection().startColumn }
      : p;
    this.drag = true;
    this.actions.selectBetween(this.anchor, p);
    e.preventDefault();
  }
  keydown(e: KeyboardEvent) {
    if (e.defaultPrevented) return;
    if (e.isComposing || this.editor.isComposing || e.keyCode === 229) return;
    const dataJumpKey =
      !this.editor.isEditing &&
      (e.ctrlKey || e.metaKey) &&
      !e.altKey &&
      ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key);
    if (!dataJumpKey && !['Control', 'Meta', 'Shift'].includes(e.key))
      this.actions.cancelDataJump();
    if (e.key === 'Escape' && this.editor.isPending) {
      e.preventDefault();
      this.editor.cancelPending();
      return;
    }
    if (this.editor.isEditing) {
      if (
        e.key === 'ArrowDown' &&
        e.altKey &&
        !e.shiftKey &&
        !e.ctrlKey &&
        !e.metaKey &&
        this.editor.canPickReference() &&
        this.actions.startKeyboardReference()
      ) {
        e.preventDefault();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        const origin = this.editor.getOrigin();
        this.editor.cancel();
        if (origin) this.actions.select(origin);
        this.actions.focus();
      } else if (e.key === 'Enter' && e.altKey && !e.ctrlKey && !e.metaKey) {
        e.preventDefault();
        this.editor.element.setRangeText(
          '\n',
          this.editor.element.selectionStart,
          this.editor.element.selectionEnd,
          'end',
        );
        this.editor.element.dispatchEvent(new Event('input', { bubbles: true }));
      } else if ((e.key === 'Enter' && !e.altKey) || e.key === 'Tab') {
        e.preventDefault();
        const origin = this.editor.getOrigin();
        if (this.editor.finish()) {
          if (origin) this.actions.select(origin);
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
    if (mod && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'd') {
      e.preventDefault();
      this.actions.run(() => {
        this.requireStableSelection();
        return this.actions.range().fillDown();
      });
      return;
    }
    if (mod && !e.altKey && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      this.actions.run(() => (e.shiftKey ? this.book.redo() : this.book.undo()));
      return;
    }
    if (mod && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      this.actions.run(() => this.book.redo());
      return;
    }
    if (mod && !e.altKey && !e.shiftKey && e.key.toLowerCase() === 'a') {
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
    if (mod && !e.altKey && moves[e.key] && e.key !== 'Tab') {
      e.preventDefault();
      this.actions.dataJump(...moves[e.key], e.shiftKey);
      return;
    }
    if (moves[e.key] && !mod && !e.altKey) {
      e.preventDefault();
      this.actions.move(...moves[e.key], e.key !== 'Tab' && e.shiftKey);
      return;
    }
    if (!mod && !e.altKey && (e.key === 'PageUp' || e.key === 'PageDown')) {
      e.preventDefault();
      this.actions.page(e.key === 'PageUp' ? -1 : 1, e.shiftKey);
      return;
    }
    if (mod && !e.altKey && e.key === 'End') {
      e.preventDefault();
      this.actions.end(e.shiftKey);
      return;
    }
    if (!mod && !e.altKey && (e.key === 'Enter' || e.key === 'F2')) {
      e.preventDefault();
      this.editor.start();
      return;
    }
    if (!mod && !e.altKey && !e.shiftKey && (e.key === 'Delete' || e.key === 'Backspace')) {
      e.preventDefault();
      if (!this.actions.readOnly())
        this.actions.run(() => {
          this.requireStableSelection();
          return this.actions.range().clear();
        });
      return;
    }
    if (!e.altKey && e.key === 'Home') {
      e.preventDefault();
      const sheet = this.book.getSheetById(this.actions.selection().sheetId)!;
      this.actions.move(mod ? -sheet.rowCount : 0, -sheet.columnCount, e.shiftKey);
      return;
    }
    if (e.key.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      this.editor.start(e.key);
    }
  }
  private requireStableSelection() {
    if (this.book.persistenceState.pending)
      throw new OpenSheetError('BUSY', 'Wait for saving to finish, then retry the action.');
  }
}
