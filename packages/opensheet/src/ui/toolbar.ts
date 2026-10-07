import { OpenSheetError, type CellStyle, type Selection, type Workbook } from '@opensheetjs/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDown,
  Bold,
  Columns3,
  Ellipsis,
  Italic,
  Merge,
  PanelTop,
  Redo2,
  Rows3,
  Split,
  Square,
  Underline,
  Undo2,
  WrapText,
  type IconNode,
} from 'lucide';
import { button } from './button.js';
let toolbarMenuId = 0;
export interface ToolbarActions {
  getWorkbook(): Workbook;
  getSelection(): Selection | null;
  run(action: () => unknown): void;
}
export class Toolbar {
  readonly element: HTMLDivElement;
  private items: HTMLDivElement;
  private moreButton: HTMLButtonElement;
  private moreMenu: HTMLDivElement;
  private pluginToolbar: HTMLDivElement;
  private cleanups: Array<() => void> = [];
  private disposed = false;
  private frame = 0;
  private abort = new AbortController();
  constructor(
    private actions: ToolbarActions,
    visible: boolean,
  ) {
    this.element = document.createElement('div');
    this.element.className = 'os-toolbar';
    this.element.setAttribute('role', 'toolbar');
    this.element.setAttribute('aria-label', 'Spreadsheet tools');
    this.items = document.createElement('div');
    this.items.className = 'os-toolbar-items';
    const more = document.createElement('div');
    more.className = 'os-toolbar-more';
    this.moreButton = this.button(
      '',
      () => this.setMoreOpen(this.moreMenu.hidden),
      'More tools',
      Ellipsis,
    );
    this.moreButton.setAttribute('aria-haspopup', 'true');
    this.moreButton.setAttribute('aria-expanded', 'false');
    this.moreButton.hidden = true;
    this.moreMenu = document.createElement('div');
    this.moreMenu.className = 'os-toolbar-menu';
    this.moreMenu.id = `os-toolbar-menu-${++toolbarMenuId}`;
    this.moreMenu.setAttribute('aria-label', 'More tools');
    this.moreMenu.hidden = true;
    this.moreMenu.addEventListener(
      'click',
      (event) => {
        if ((event.target as Element).closest('button')) this.setMoreOpen(false);
      },
      { signal: this.abort.signal },
    );
    this.moreButton.setAttribute('aria-controls', this.moreMenu.id);
    more.append(this.moreButton, this.moreMenu);
    this.element.append(this.items, more);
    this.pluginToolbar = document.createElement('div');
    this.pluginToolbar.className = 'os-plugin-toolbar';
    this.buildToolbar();
    this.items.append(this.pluginToolbar);
    this.element.hidden = !visible;
    const resize = new ResizeObserver(() => this.layoutToolbar());
    resize.observe(this.element);
    this.cleanups.push(() => resize.disconnect());
    const dismissMore = (event: PointerEvent) => {
      if (!more.contains(event.target as Node)) this.setMoreOpen(false);
    };
    const escapeMore = (event: KeyboardEvent) => {
      if (event.key === 'Escape' && !this.moreMenu.hidden) {
        this.setMoreOpen(false);
        this.moreButton.focus();
      }
    };
    document.addEventListener('pointerdown', dismissMore);
    document.addEventListener('keydown', escapeMore);
    this.cleanups.push(() => document.removeEventListener('pointerdown', dismissMore));
    this.cleanups.push(() => document.removeEventListener('keydown', escapeMore));
    this.frame = requestAnimationFrame(() => {
      this.frame = 0;
      this.layoutToolbar();
    });
  }
  private button(label: string, action: () => void, title = label, icon?: IconNode) {
    return button(label, () => this.actions.run(action), title, icon, this.abort.signal);
  }
  private setMoreOpen(open: boolean) {
    this.moreMenu.hidden = !open;
    this.moreButton.setAttribute('aria-expanded', String(open));
  }
  private layoutToolbar() {
    if (this.disposed || this.element.hidden || !this.element.isConnected) return;
    this.setMoreOpen(false);
    this.items.append(...this.moreMenu.children);
    this.moreButton.hidden = true;
    if (this.items.scrollWidth <= this.items.clientWidth) return;
    this.moreButton.hidden = false;
    while (this.items.scrollWidth > this.items.clientWidth) {
      const last = this.items.lastElementChild;
      if (!last) break;
      this.moreMenu.prepend(last);
    }
  }
  private buildToolbar() {
    const group = (...buttons: HTMLElement[]) => {
      const g = document.createElement('div');
      g.className = 'os-tool-group';
      g.append(...buttons);
      this.items.append(g);
    };
    const range = () => {
      const s = this.actions.getSelection();
      if (!s) throw new Error('No selection');
      return this.actions.getWorkbook().getSheetById(s.sheetId)!.range(s);
    };
    const style = (s: CellStyle) => range().setStyle(s);
    group(
      this.button('', () => this.actions.getWorkbook().undo(), 'Undo', Undo2),
      this.button('', () => this.actions.getWorkbook().redo(), 'Redo', Redo2),
    );
    const fmt = document.createElement('select');
    fmt.setAttribute('aria-label', 'Number format');
    for (const [value, label] of [
      ['General', 'Automatic'],
      ['0.00', 'Number'],
      ['$#,##0.00', 'Currency'],
      ['0.0%', 'Percent'],
      ['yyyy-mm-dd', 'Date'],
    ]) {
      const o = document.createElement('option');
      o.value = value;
      o.textContent = label;
      fmt.append(o);
    }
    fmt.addEventListener(
      'change',
      () => this.actions.run(() => style({ numberFormat: fmt.value })),
      { signal: this.abort.signal },
    );
    group(fmt);
    group(
      this.button(
        '',
        () => {
          const s = this.actions.getSelection()!,
            c = this.actions.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ bold: !this.actions.getWorkbook().getStyle(c?.styleId).bold });
        },
        'Bold',
        Bold,
      ),
      this.button(
        '',
        () => {
          const s = this.actions.getSelection()!,
            c = this.actions.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ italic: !this.actions.getWorkbook().getStyle(c?.styleId).italic });
        },
        'Italic',
        Italic,
      ),
      this.button(
        '',
        () => {
          const s = this.actions.getSelection()!,
            c = this.actions.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ underline: !this.actions.getWorkbook().getStyle(c?.styleId).underline });
        },
        'Underline',
        Underline,
      ),
    );
    const color = document.createElement('input');
    color.type = 'color';
    color.value = '#e1efe7';
    color.title = 'Cell background';
    color.setAttribute('aria-label', 'Cell background');
    color.addEventListener(
      'input',
      () => this.actions.run(() => style({ background: color.value })),
      { signal: this.abort.signal },
    );
    group(
      color,
      this.button('', () => style({ align: 'left' }), 'Align left', AlignLeft),
      this.button('', () => style({ align: 'center' }), 'Align center', AlignCenter),
      this.button('', () => style({ align: 'right' }), 'Align right', AlignRight),
    );
    group(
      this.button('Merge', () => range().merge(), 'Merge', Merge),
      this.button('Unmerge', () => range().unmerge(), 'Unmerge', Split),
      this.button('Borders', () => style({ border: true }), 'Borders', Square),
      this.button(
        'Wrap',
        () => {
          const s = this.actions.getSelection()!,
            c = this.actions.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ wrap: !this.actions.getWorkbook().getStyle(c?.styleId).wrap });
        },
        'Wrap text',
        WrapText,
      ),
      this.button('Fill', () => range().fillDown(), 'Fill down', ArrowDown),
    );
    group(
      this.button(
        'Row',
        () => {
          const s = this.actions.getSelection()!;
          this.actions.getWorkbook().getSheetById(s.sheetId)!.insertRows(s.startRow);
        },
        'Insert row',
        Rows3,
      ),
      this.button(
        'Column',
        () => {
          const s = this.actions.getSelection()!;
          this.actions.getWorkbook().getSheetById(s.sheetId)!.insertColumns(s.startColumn);
        },
        'Insert column',
        Columns3,
      ),
    );
    group(
      this.button(
        'Freeze',
        () => {
          const s = this.actions.getSelection()!,
            book = this.actions.getWorkbook(),
            current = book.sheetData(s.sheetId).freeze;
          book.getSheetById(s.sheetId)!.setFreeze(current.rows ? 0 : 1, 0);
        },
        'Freeze first row',
        PanelTop,
      ),
    );
  }
  add(action: { id: string; label: string; run(): void }) {
    if (this.pluginToolbar.querySelector(`[data-action="${CSS.escape(action.id)}"]`))
      throw new OpenSheetError('PLUGIN_CONFLICT', 'Duplicate toolbar action');
    const abort = new AbortController();
    const b = button(
      action.label,
      () => this.actions.run(action.run),
      action.label,
      undefined,
      abort.signal,
    );
    b.dataset.action = action.id;
    this.pluginToolbar.append(b);
    this.layoutToolbar();
    return () => {
      abort.abort();
      b.remove();
      this.layoutToolbar();
    };
  }
  dispose() {
    this.disposed = true;
    cancelAnimationFrame(this.frame);
    this.abort.abort();
    this.cleanups.forEach((fn) => fn());
  }
}
