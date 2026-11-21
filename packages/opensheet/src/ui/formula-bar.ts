import {
  address,
  parseRange,
  rangeAddress,
  OpenSheetError,
  LIMITS,
  parseInput,
  formulaErrorExplanation,
  type CellInput,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
import {
  FormulaAssist,
  canInsertFormulaReference,
  resolveFormulaReferenceSpans,
  formulaReferenceColor,
} from '@opensheetjs/renderer';
let nextErrorId = 0;
export interface FormulaActions {
  book(): Workbook;
  selection(): Selection | null;
  select(selection: Selection): void;
  move(rows: number, columns: number, extend: boolean): void;
  references(selections: Selection[]): void;
  draft(dirty: boolean): void;
  focus(): void;
  run(action: () => unknown): void;
}
export class FormulaBar {
  readonly element: HTMLDivElement;
  private namebox: HTMLInputElement;
  private formula: HTMLInputElement;
  private referenceMirror = document.createElement('div');
  private referenceResize?: ResizeObserver;
  private abort = new AbortController();
  private assist: FormulaAssist;
  private errorButton = document.createElement('button');
  private errorDetails = document.createElement('div');
  private errorIdentity = '';
  private editTarget?: Selection;
  private editBook?: Workbook;
  private editIdentity?: { rowId: number; columnId: number; input?: CellInput | null };
  private editRevision = 0;
  private reference?: { start: number; end: number };
  private composing = false;
  private baseline = '';
  private editingDirty = false;
  private editGeneration = 0;
  private pendingDrafts = 0;
  private interactingRecovery = false;
  private recovery = document.createElement('section');
  private keyboardPick?: {
    value: string;
    start: number;
    end: number;
    selection: Selection;
    anchor: { row: number; column: number };
    focus: { row: number; column: number };
  };
  constructor(private actions: FormulaActions) {
    this.namebox = document.createElement('input');
    this.namebox.className = 'os-namebox';
    this.namebox.setAttribute('aria-label', 'Cell address');
    this.namebox.value = 'A1';
    this.formula = document.createElement('input');
    this.formula.className = 'os-formula-input';
    this.formula.setAttribute('aria-label', 'Formula bar');
    this.formula.placeholder = 'Enter a value or formula, e.g. =SUM(D2:D10)';
    this.element = document.createElement('div');
    this.element.className = 'os-formula-bar';
    const fx = document.createElement('span');
    fx.textContent = 'ƒx';
    const field = document.createElement('div');
    field.className = 'os-formula-field';
    this.referenceMirror.className = 'os-formula-reference-mirror';
    this.referenceMirror.setAttribute('aria-hidden', 'true');
    field.append(this.formula, this.referenceMirror);
    this.element.append(this.namebox, fx, field);
    this.recovery.className = 'os-draft-recovery os-formula-draft-recovery';
    this.recovery.setAttribute('aria-label', 'Unconfirmed formula bar drafts');
    this.recovery.hidden = true;
    this.element.append(this.recovery);
    this.recovery.addEventListener(
      'pointerdown',
      () => {
        this.interactingRecovery = true;
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'scroll',
      () => {
        this.syncReferenceMirror();
      },
      { signal: this.abort.signal },
    );
    this.errorButton.type = 'button';
    this.errorButton.className = 'os-formula-error-button';
    this.errorButton.hidden = true;
    this.errorButton.setAttribute('aria-expanded', 'false');
    this.errorDetails.className = 'os-formula-error-details';
    this.errorDetails.id = `os-formula-error-${++nextErrorId}`;
    this.errorButton.setAttribute('aria-controls', this.errorDetails.id);
    this.errorDetails.hidden = true;
    this.errorDetails.setAttribute('role', 'status');
    this.element.append(this.errorButton, this.errorDetails);
    this.errorButton.addEventListener(
      'click',
      () => {
        this.errorDetails.hidden = !this.errorDetails.hidden;
        this.errorButton.setAttribute('aria-expanded', String(!this.errorDetails.hidden));
      },
      { signal: this.abort.signal },
    );
    this.errorButton.addEventListener(
      'keydown',
      (event) => {
        if (event.key === 'Escape' && !this.errorDetails.hidden) {
          event.preventDefault();
          this.errorDetails.hidden = true;
          this.errorButton.setAttribute('aria-expanded', 'false');
        }
      },
      { signal: this.abort.signal },
    );
    this.assist = new FormulaAssist(this.formula, this.abort.signal);
    this.element.append(this.assist.element);
    this.formula.addEventListener(
      'focus',
      () => {
        this.editGeneration++;
        if (!this.editingDirty || !this.editBook) {
          this.baseline = this.formula.value;
          const selection = this.actions.selection();
          this.editTarget = selection ? { ...selection } : undefined;
          this.editBook = selection ? this.actions.book() : undefined;
          const sheet = selection && this.editBook?.sheetData(selection.sheetId);
          const loaded =
            selection &&
            this.editBook?.isLoaded(selection.sheetId, selection.startRow, selection.startColumn);
          this.editIdentity =
            selection && sheet
              ? {
                  rowId: sheet.rowOrder[selection.startRow],
                  columnId: sheet.columnOrder[selection.startColumn],
                  ...(loaded
                    ? {
                        input: structuredClone(
                          this.editBook?.peekCell(
                            selection.sheetId,
                            selection.startRow,
                            selection.startColumn,
                          )?.input ?? null,
                        ),
                      }
                    : {}),
                }
              : undefined;
          this.editRevision = this.editBook?.revision ?? 0;
          if (selection && this.editBook && this.editIdentity && !loaded) {
            const book = this.editBook;
            const identity = this.editIdentity;
            const revision = this.editRevision;
            void book.getCell(selection.sheetId, selection.startRow, selection.startColumn).then(
              (cell) => {
                if (
                  this.editBook === book &&
                  this.editIdentity === identity &&
                  book.revision === revision
                )
                  identity.input = structuredClone(cell?.input ?? null);
              },
              () => {
                // Keep the revision precondition if the original value could not be read.
              },
            );
          }
        }
        this.refreshReferences();
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'blur',
      (event) => {
        this.editGeneration++;
        this.composing = false;
        if (
          this.interactingRecovery ||
          (event.relatedTarget instanceof Node && this.recovery.contains(event.relatedTarget))
        )
          return;
        this.setEditingDirty(false);
        this.editTarget = undefined;
        this.editBook = undefined;
        this.editIdentity = undefined;
        this.reference = undefined;
        this.keyboardPick = undefined;
        this.assist.setPicking(false);
        this.actions.references([]);
        this.referenceMirror.replaceChildren();
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'compositionstart',
      () => {
        this.composing = true;
        this.referenceMirror.style.visibility = 'hidden';
        this.actions.references([]);
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'compositionend',
      () => {
        this.composing = false;
        this.refreshReferences();
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'input',
      () => {
        this.editGeneration++;
        this.setEditingDirty(this.formula.value !== this.baseline);
        if (this.formula.value !== this.baseline) this.clearError();
        if (!this.composing) this.refreshReferences();
      },
      { signal: this.abort.signal },
    );
    this.namebox.addEventListener(
      'keydown',
      (e) => {
        if (e.isComposing || e.keyCode === 229) return;
        if (e.key === 'Enter' && (e.altKey || e.ctrlKey || e.metaKey)) return;
        if (e.key === 'Escape') {
          e.preventDefault();
          const selection = this.actions.selection();
          if (selection) this.namebox.value = rangeAddress(selection);
          this.actions.focus();
        } else if (e.key === 'Enter') {
          e.preventDefault();
          const input = this.namebox.value.trim();
          this.actions.run(() => {
            const book = this.actions.book();
            const selection = this.actions.selection();
            if (!selection) throw new OpenSheetError('INVALID_RANGE', 'No active selection');
            const qualified = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(input);
            const sheetName = qualified
              ? (qualified[1]?.replace(/''/g, "'") ?? qualified[2])
              : undefined;
            const sheet = sheetName
              ? book.getSheetByName(sheetName)
              : book.getSheetById(selection.sheetId);
            if (!sheet) throw new OpenSheetError('INVALID_RANGE', 'Unknown worksheet');
            const range = parseRange(qualified ? qualified[3] : input);
            sheet.range(range); // Validate against the target sheet before switching sheets.
            this.actions.select({
              ...range,
              sheetId: sheet.id,
            });
            const selected = this.actions.selection();
            if (selected) this.namebox.value = rangeAddress(selected);
            this.actions.focus();
          });
        }
      },
      { signal: this.abort.signal },
    );
    this.namebox.addEventListener(
      'blur',
      () => {
        const selection = this.actions.selection();
        if (selection) this.namebox.value = rangeAddress(selection);
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'keydown',
      (e) => {
        if (
          e.defaultPrevented ||
          this.composing ||
          e.isComposing ||
          e.keyCode === 229 ||
          this.formula.readOnly
        )
          return;
        if (e.key === 'Enter' && (e.altKey || e.ctrlKey || e.metaKey)) return;
        if (this.keyboardPick) {
          if (this.handleKeyboardReference(e)) return;
          this.endReference();
          this.keyboardPick = undefined;
          this.assist.setPicking(false);
        }
        if (
          e.altKey &&
          !e.shiftKey &&
          !e.ctrlKey &&
          !e.metaKey &&
          e.key === 'ArrowDown' &&
          this.startKeyboardReference()
        ) {
          e.preventDefault();
          return;
        }
        if (e.key === 'Enter') {
          const target = this.editTarget ? { ...this.editTarget } : this.actions.selection();
          const identity = this.editIdentity ? { ...this.editIdentity } : undefined;
          const editingBook = this.editBook;
          const revision = this.editRevision;
          const value = this.formula.value;
          const unchanged = value === this.baseline;
          const generation = this.editGeneration;
          this.actions.run(async () => {
            const s = target;
            if (!s || !identity || editingBook !== this.actions.book())
              throw new OpenSheetError('REVISION_CONFLICT', 'Select the cell again before editing');
            if (unchanged) {
              if (!this.abort.signal.aborted && this.editGeneration === generation) {
                if (this.actions.book().revision === revision) this.actions.select(s);
                this.actions.focus();
              }
              return;
            }
            if (value.length > LIMITS.text + 1)
              throw new OpenSheetError('LIMIT_EXCEEDED', 'Cell input is too long');
            const book = this.actions.book();
            const sheet = book.getSheetById(s.sheetId)!;
            const location = `${sheet.name}!${address(s.startRow, s.startColumn)}`;
            if (this.pendingDrafts + this.recovery.childElementCount >= 32)
              throw new OpenSheetError(
                'LIMIT_EXCEEDED',
                'Resolve unconfirmed drafts before submitting more',
              );
            this.pendingDrafts++;
            try {
              const command = {
                type: 'core.cells.set',
                payload: {
                  sheetId: s.sheetId,
                  cells: [
                    {
                      row: s.startRow,
                      column: s.startColumn,
                      input: parseInput(value),
                      expected: identity,
                    },
                  ],
                },
              };
              await book.execute(
                identity.input === undefined
                  ? {
                      ...command,
                      protocolVersion: 1,
                      commandId: crypto.randomUUID(),
                      workbookId: book.id,
                      baseRevision: revision,
                    }
                  : command,
              );
            } catch (error) {
              if (!this.abort.signal.aborted) this.retainDraft(location, value);
              throw error;
            } finally {
              this.pendingDrafts--;
            }
            if (
              !this.abort.signal.aborted &&
              this.editGeneration === generation &&
              this.actions.book() === book
            ) {
              this.actions.select(s);
              this.actions.focus();
            }
          });
        }
        if (e.key === 'Escape') {
          const target = this.editTarget;
          this.formula.blur();
          if (target) this.actions.select(target);
          this.update(this.actions.book(), this.actions.selection());
          this.actions.focus();
        }
      },
      { signal: this.abort.signal },
    );
    try {
      document.addEventListener(
        'pointerup',
        () => {
          this.interactingRecovery = false;
        },
        { signal: this.abort.signal },
      );
      document.addEventListener(
        'pointercancel',
        () => {
          this.interactingRecovery = false;
        },
        { signal: this.abort.signal },
      );
      if (typeof ResizeObserver !== 'undefined') {
        this.referenceResize = new ResizeObserver(() => this.syncReferenceMirror());
        this.referenceResize.observe(this.formula);
      }
    } catch (error) {
      try {
        this.dispose();
      } catch (cleanupError) {
        console.error(cleanupError);
      }
      throw error;
    }
  }
  private setEditingDirty(dirty: boolean) {
    this.editingDirty = dirty;
    this.actions.draft(dirty || this.recovery.childElementCount > 0);
  }
  private retainDraft(location: string, value: string) {
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
      this.setEditingDirty(this.editingDirty);
    };
    entry.append(label, dismiss);
    this.recovery.append(entry);
    this.recovery.hidden = false;
    this.setEditingDirty(this.editingDirty);
  }
  setReadOnly(read: boolean) {
    this.formula.readOnly = read;
    if (read) {
      this.assist.hide();
      if (document.activeElement === this.formula) {
        const editingBook = this.editBook;
        this.formula.blur();
        const book = this.actions.book();
        if (editingBook === book) this.update(book, this.actions.selection());
      }
    }
  }
  /** True when a grid reference can be inserted at the current caret. */
  canPickReference(): boolean {
    if (
      document.activeElement !== this.formula ||
      this.composing ||
      this.formula.readOnly ||
      !this.editTarget ||
      !this.formula.value.startsWith('=')
    )
      return false;
    return canInsertFormulaReference(this.formula.value, this.formula.selectionStart ?? 0);
  }
  beginReference(): boolean {
    if (this.keyboardPick || !this.canPickReference()) return false;
    const start = this.formula.selectionStart ?? 0;
    this.reference = { start, end: this.formula.selectionEnd ?? start };
    return true;
  }
  updateReference(selection: Selection) {
    if (!this.reference || !this.editTarget) return;
    const sheet = this.actions.book().getSheetById(selection.sheetId);
    if (!sheet) return;
    const prefix =
      selection.sheetId === this.editTarget.sheetId ? '' : `'${sheet.name.replace(/'/g, "''")}'!`;
    const text = prefix + rangeAddress(selection);
    this.formula.setRangeText(text, this.reference.start, this.reference.end, 'end');
    this.reference.end = this.reference.start + text.length;
    this.formula.dispatchEvent(new Event('input', { bubbles: true }));
  }
  endReference() {
    this.reference = undefined;
  }
  adoptDraft(
    value: string,
    start: number,
    end: number,
    identity: { rowId: number; columnId: number; input?: CellInput | null },
    revision: number,
  ) {
    const baseline = this.formula.value;
    this.formula.value = value;
    this.formula.focus();
    this.editIdentity = { ...identity };
    this.editRevision = revision;
    this.baseline = baseline;
    this.formula.setSelectionRange(start, end);
    this.formula.dispatchEvent(new Event('input', { bubbles: true }));
  }
  private refreshReferences() {
    if (!this.editTarget || !this.editBook) return;
    const source = this.formula.value;
    const references = resolveFormulaReferenceSpans(this.editBook, this.editTarget.sheetId, source);
    this.actions.references(references.map(({ selection }) => selection));
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
    this.referenceMirror.replaceChildren(...nodes);
    this.syncReferenceMirror();
  }
  private syncReferenceMirror() {
    // WebKit may scroll the native input internally while reporting scrollLeft=0.
    // Suppress underlines when text overflows rather than show them on wrong tokens.
    this.referenceMirror.style.visibility =
      this.composing || this.formula.scrollWidth > this.formula.clientWidth + 1 ? 'hidden' : '';
    this.referenceMirror.scrollLeft = this.formula.scrollLeft;
  }
  startKeyboardReference(): boolean {
    const selection = this.actions.selection();
    if (!selection || !this.beginReference()) return false;
    this.keyboardPick = {
      value: this.formula.value,
      start: this.formula.selectionStart ?? 0,
      end: this.formula.selectionEnd ?? 0,
      selection: { ...selection },
      anchor: { row: selection.startRow, column: selection.startColumn },
      focus: { row: selection.startRow, column: selection.startColumn },
    };
    this.assist.setPicking(true);
    this.updateReference(selection);
    this.announceReference(selection);
    return true;
  }
  private announceReference(selection: Selection) {
    const sheet = this.actions.book().getSheetById(selection.sheetId);
    this.assist.announceReference(
      `${sheet?.name ?? ''} ${rangeAddress(selection)}. Arrows move, Shift extends, Page Up or Page Down changes sheet, Enter accepts, Escape cancels.`,
    );
  }
  private handleKeyboardReference(event: KeyboardEvent): boolean {
    const pick = this.keyboardPick!;
    if (['Shift', 'Control', 'Alt', 'Meta'].includes(event.key)) return true;
    if (event.key === 'Escape') {
      event.preventDefault();
      this.endReference();
      this.keyboardPick = undefined;
      this.formula.value = pick.value;
      this.formula.setSelectionRange(pick.start, pick.end);
      this.actions.select(pick.selection);
      this.assist.setPicking(false);
      this.formula.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    }
    if (event.key === 'Enter') {
      event.preventDefault();
      this.endReference();
      this.keyboardPick = undefined;
      this.assist.setPicking(false);
      this.formula.dispatchEvent(new Event('input', { bubbles: true }));
      return true;
    }
    const moves: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    };
    const movement = moves[event.key];
    if (movement && !event.altKey && !event.ctrlKey && !event.metaKey) {
      event.preventDefault();
      if (event.shiftKey) {
        const current = this.actions.selection()!;
        this.actions.select({
          sheetId: current.sheetId,
          startRow: pick.focus.row,
          endRow: pick.focus.row + 1,
          startColumn: pick.focus.column,
          endColumn: pick.focus.column + 1,
        });
        this.actions.move(movement[0], movement[1], false);
        const moved = this.actions.selection()!;
        pick.focus = {
          row: movement[0] > 0 ? moved.endRow - 1 : moved.startRow,
          column: movement[1] > 0 ? moved.endColumn - 1 : moved.startColumn,
        };
        this.actions.select({
          sheetId: moved.sheetId,
          startRow: Math.min(pick.anchor.row, pick.focus.row),
          endRow: Math.max(pick.anchor.row, pick.focus.row) + 1,
          startColumn: Math.min(pick.anchor.column, pick.focus.column),
          endColumn: Math.max(pick.anchor.column, pick.focus.column) + 1,
        });
      } else {
        this.actions.move(movement[0], movement[1], false);
        const moved = this.actions.selection()!;
        pick.anchor = { row: moved.startRow, column: moved.startColumn };
        pick.focus = { ...pick.anchor };
      }
    } else if (
      (event.key === 'PageUp' || event.key === 'PageDown') &&
      !event.altKey &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.shiftKey
    ) {
      event.preventDefault();
      const book = this.actions.book();
      const sheets = book.getSheets().filter((sheet) => !book.sheetData(sheet.id).hidden);
      const current = sheets.findIndex((sheet) => sheet.id === this.actions.selection()?.sheetId);
      const next =
        sheets[
          Math.max(0, Math.min(sheets.length - 1, current + (event.key === 'PageDown' ? 1 : -1)))
        ];
      if (next)
        this.actions.select({
          sheetId: next.id,
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 1,
        });
      pick.anchor = { row: 0, column: 0 };
      pick.focus = { ...pick.anchor };
    } else return false;
    const selection = this.actions.selection();
    if (selection) {
      this.updateReference(selection);
      this.announceReference(selection);
    }
    return true;
  }
  update(book: Workbook, s: Selection | null) {
    if (this.editBook && this.editBook !== book) this.formula.blur();
    if (!s) {
      this.clearError();
      return;
    }
    if (document.activeElement !== this.namebox) this.namebox.value = rangeAddress(s);
    const c = book.peekCell(s.sheetId, s.startRow, s.startColumn);
    const value = book.peekValue(s.sheetId, s.startRow, s.startColumn);
    const code =
      c?.input.type === 'formula' && value && typeof value === 'object' ? value.error : '';
    const explanation = code ? formulaErrorExplanation(code) : undefined;
    const identity = `${book.id}/${s.sheetId}/${s.startRow}/${s.startColumn}/${code}`;
    if (
      !explanation ||
      (document.activeElement === this.formula && this.formula.value !== this.baseline)
    ) {
      this.clearError();
    } else {
      if (this.errorIdentity !== identity) {
        this.errorDetails.hidden = true;
        this.errorButton.setAttribute('aria-expanded', 'false');
      }
      this.errorIdentity = identity;
      this.errorButton.hidden = false;
      this.errorButton.textContent = code;
      this.errorButton.setAttribute('aria-label', `Explain formula error ${code}`);
      this.errorDetails.textContent = `${code}: ${explanation} The error may originate in a referenced cell.`;
    }
    if (document.activeElement !== this.formula && !this.editingDirty)
      this.formula.value =
        c?.input.type === 'formula'
          ? '=' + c.input.expression
          : c?.input.type === 'error'
            ? c.input.code
            : c && 'value' in c.input
              ? String(c.input.value)
              : '';
  }
  private clearError() {
    const restoreFocus = document.activeElement === this.errorButton;
    this.errorIdentity = '';
    this.errorButton.hidden = true;
    this.errorDetails.hidden = true;
    this.errorButton.setAttribute('aria-expanded', 'false');
    if (restoreFocus) this.actions.focus();
  }
  dispose() {
    try {
      this.referenceResize?.disconnect();
    } finally {
      this.abort.abort();
    }
  }
}
