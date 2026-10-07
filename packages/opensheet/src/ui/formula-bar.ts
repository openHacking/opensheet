import {
  address,
  parseRange,
  rangeAddress,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
export interface FormulaActions {
  book(): Workbook;
  selection(): Selection | null;
  select(selection: Selection): void;
  focus(): void;
  run(action: () => unknown): void;
}
export class FormulaBar {
  readonly element: HTMLDivElement;
  private namebox: HTMLInputElement;
  private formula: HTMLInputElement;
  private abort = new AbortController();
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
    this.element.append(this.namebox, fx, this.formula);
    this.namebox.addEventListener(
      'keydown',
      (e) => {
        if (e.key === 'Enter')
          this.actions.run(() => {
            this.actions.select({
              ...parseRange(this.namebox.value),
              sheetId: this.actions.selection()!.sheetId,
            });
            this.actions.focus();
          });
      },
      { signal: this.abort.signal },
    );
    this.formula.addEventListener(
      'keydown',
      (e) => {
        if (e.key === 'Enter')
          this.actions.run(() => {
            const s = this.actions.selection()!;
            this.actions
              .book()
              .getSheetById(s.sheetId)!
              .range(address(s.startRow, s.startColumn))
              .setInput(this.formula.value);
            this.actions.focus();
          });
        if (e.key === 'Escape') {
          this.update(this.actions.book(), this.actions.selection());
          this.actions.focus();
        }
      },
      { signal: this.abort.signal },
    );
  }
  setReadOnly(read: boolean) {
    this.formula.readOnly = read;
  }
  update(book: Workbook, s: Selection | null) {
    if (!s) return;
    this.namebox.value = rangeAddress(s);
    const c = book.getCell(s.sheetId, s.startRow, s.startColumn);
    if (document.activeElement !== this.formula)
      this.formula.value =
        c?.input.type === 'formula'
          ? '=' + c.input.expression
          : c?.input.type === 'error'
            ? c.input.code
            : c && 'value' in c.input
              ? String(c.input.value)
              : '';
  }
  dispose() {
    this.abort.abort();
  }
}
