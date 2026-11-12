import { formulaHelp, type FormulaHelp } from '@opensheetjs/core';
import { cycleFormulaReference } from './formula-references.js';
let nextId = 0;

export class FormulaAssist {
  readonly element = document.createElement('div');
  private list = document.createElement('div');
  private hint = document.createElement('div');
  private help: FormulaHelp = {};
  private selected = 0;
  private composing = false;
  private picking = false;
  constructor(
    private input: HTMLInputElement | HTMLTextAreaElement,
    signal: AbortSignal,
    private mode: 'bar' | 'cell' = 'bar',
  ) {
    const id = `os-formula-help-${++nextId}`;
    this.element.className = 'os-formula-assist';
    this.list.id = id;
    this.list.setAttribute('role', 'listbox');
    this.list.setAttribute('aria-label', 'Formula functions');
    this.hint.id = `${id}-hint`;
    this.hint.setAttribute('role', 'status');
    this.hint.setAttribute('aria-live', 'polite');
    this.hint.setAttribute('aria-atomic', 'true');
    this.element.append(this.list, this.hint);
    if (mode === 'bar') input.setAttribute('role', 'combobox');
    input.setAttribute('aria-autocomplete', 'list');
    if (mode === 'bar') {
      input.setAttribute('aria-keyshortcuts', 'Alt+ArrowDown F4');
      input.title = 'Alt+Down: select a cell reference; F4: change reference anchoring';
    } else {
      input.setAttribute('aria-keyshortcuts', 'F4');
    }
    input.setAttribute('aria-controls', id);
    input.setAttribute('aria-describedby', this.hint.id);
    this.hide();
    input.addEventListener('input', () => this.refresh(), { signal });
    input.addEventListener('focus', () => this.refresh(), { signal });
    input.addEventListener('click', () => this.refresh(), { signal });
    input.addEventListener(
      'keyup',
      (e) => {
        if (['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes((e as KeyboardEvent).key))
          this.refresh();
      },
      { signal },
    );
    input.addEventListener(
      'blur',
      () => {
        this.composing = false;
        this.hide();
      },
      { signal },
    );
    input.addEventListener(
      'compositionstart',
      () => {
        this.composing = true;
        this.hide();
      },
      { signal },
    );
    input.addEventListener(
      'compositionend',
      () => {
        this.composing = false;
        this.refresh();
      },
      { signal },
    );
    input.addEventListener(
      'keydown',
      (event) => {
        const e = event as KeyboardEvent;
        if (
          this.picking ||
          this.composing ||
          e.isComposing ||
          e.keyCode === 229 ||
          input.readOnly ||
          e.altKey ||
          e.ctrlKey ||
          e.metaKey
        )
          return;
        if (e.key === 'F4' && !e.shiftKey) {
          const start = input.selectionStart ?? 0;
          const end = input.selectionEnd ?? start;
          const edit = cycleFormulaReference(input.value, start, end);
          if (edit) {
            e.preventDefault();
            input.setRangeText(edit.text, edit.start, edit.end, start === end ? 'end' : 'select');
            input.dispatchEvent(new Event('input', { bubbles: true }));
          }
          return;
        }
        const functions = this.help.completion?.functions;
        if (!functions?.length) return;
        if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
          e.preventDefault();
          this.selected =
            (this.selected + (e.key === 'ArrowDown' ? 1 : -1) + functions.length) %
            functions.length;
          this.renderOptions();
        } else if (e.key === 'Enter' || e.key === 'Tab') {
          e.preventDefault();
          this.accept();
        } else if (e.key === 'Escape') {
          e.preventDefault();
          this.hide();
        }
      },
      { signal },
    );
    this.list.addEventListener(
      'mousedown',
      (event) => {
        const option = (event.target as Element).closest<HTMLElement>('[data-function-index]');
        if (!option) return;
        event.preventDefault();
        this.selected = Number(option.dataset.functionIndex);
        this.accept();
      },
      { signal },
    );
  }
  hide() {
    this.help = {};
    this.list.replaceChildren();
    this.hint.textContent = '';
    this.element.hidden = true;
    if (this.mode === 'bar') this.input.setAttribute('aria-expanded', 'false');
    this.input.removeAttribute('aria-activedescendant');
  }
  setPicking(picking: boolean) {
    this.picking = picking;
    this.hide();
  }
  announceReference(text: string) {
    this.hint.textContent = text;
    this.element.hidden = false;
  }
  private refresh() {
    if (this.picking) return;
    if (this.composing || this.input.readOnly || document.activeElement !== this.input) {
      this.hide();
      return;
    }
    this.help = formulaHelp(this.input.value, this.input.selectionStart ?? this.input.value.length);
    this.selected = 0;
    this.renderOptions();
    const call = this.help.call;
    this.hint.textContent = call
      ? `${call.definition.signature} — Argument ${call.argument + 1}. ${call.definition.description}`
      : '';
    this.element.hidden = !this.help.completion && !call;
  }
  private renderOptions() {
    this.list.replaceChildren();
    const functions = this.help.completion?.functions ?? [];
    functions.forEach((definition, index) => {
      const option = document.createElement('div');
      option.id = `${this.list.id}-${index}`;
      option.dataset.functionIndex = String(index);
      option.setAttribute('role', 'option');
      option.setAttribute('aria-selected', String(index === this.selected));
      option.textContent = definition.signature;
      this.list.append(option);
      if (index === this.selected) this.input.setAttribute('aria-activedescendant', option.id);
    });
    if (this.mode === 'bar') this.input.setAttribute('aria-expanded', String(functions.length > 0));
    if (!functions.length) this.input.removeAttribute('aria-activedescendant');
  }
  private accept() {
    const completion = this.help.completion;
    const definition = completion?.functions[this.selected];
    if (!completion || !definition) return;
    const suffix = this.input.value.slice(completion.end);
    const text = definition.name + (suffix.startsWith('(') ? '' : '(');
    this.input.setRangeText(text, completion.start, completion.end, 'end');
    if (suffix.startsWith('(')) {
      const cursor = (this.input.selectionStart ?? 0) + 1;
      this.input.setSelectionRange(cursor, cursor);
    }
    this.input.dispatchEvent(new Event('input', { bubbles: true }));
  }
}
