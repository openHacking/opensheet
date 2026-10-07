import {
  Workbook,
  createSnapshot,
  validateSnapshot,
  commandSchemas,
  parseRange,
  address,
  rangeAddress,
  OpenSheetError,
  type WorkbookSnapshot,
  type Selection,
  type Commit,
  type CellStyle,
  type Command,
  type CommandEnvelope,
} from '@opensheetjs/core';
import { CanvasGrid } from '@opensheetjs/renderer';
import { PluginRegistry, type Plugin } from '@opensheetjs/plugin-sdk';
import {
  createElement,
  Undo2,
  Redo2,
  Bold,
  Italic,
  Underline,
  AlignLeft,
  AlignCenter,
  AlignRight,
  Merge,
  Split,
  Square,
  WrapText,
  ArrowDown,
  Plus,
  Rows3,
  Columns3,
  PanelTop,
  Ellipsis,
  type IconNode,
} from 'lucide';
export * from '@opensheetjs/core';
export { definePlugin } from '@opensheetjs/plugin-sdk';
export type { Plugin, PluginContext } from '@opensheetjs/plugin-sdk';
export interface OpenSheetOptions {
  container: string | HTMLElement;
  mode?: 'edit' | 'read';
  toolbar?: boolean;
  onError?: (error: unknown) => void;
}
type EventMap = {
  'workbook:committed': Commit;
  'selection:changed': Selection;
  'plugin:error': unknown;
  'lifecycle:disposed': undefined;
};
let toolbarMenuId = 0;
export class OpenSheet {
  readonly element: HTMLDivElement;
  readonly plugins: PluginRegistry;
  private book?: Workbook;
  private grid?: CanvasGrid;
  private toolbar: HTMLDivElement;
  private toolbarItems: HTMLDivElement;
  private moreButton: HTMLButtonElement;
  private moreMenu: HTMLDivElement;
  private pluginToolbar: HTMLDivElement;
  private formula: HTMLInputElement;
  private namebox: HTMLInputElement;
  private tabs: HTMLDivElement;
  private viewport: HTMLDivElement;
  private status: HTMLDivElement;
  private toast: HTMLDivElement;
  private stop?: () => void;
  private disposed = false;
  private pluginList: Plugin[] = [];
  private events = new Map<string, Set<(data: any) => void>>();
  private cleanups: Array<() => void> = [];
  private timer?: ReturnType<typeof setTimeout>;
  constructor(private options: OpenSheetOptions) {
    const container =
      typeof options.container === 'string'
        ? document.querySelector<HTMLElement>(options.container)
        : options.container;
    if (!container) throw new OpenSheetError('INVALID_ARGUMENT', 'Container was not found');
    this.element = document.createElement('div');
    this.element.className = 'opensheet';
    this.toolbar = document.createElement('div');
    this.toolbar.className = 'os-toolbar';
    this.toolbar.setAttribute('role', 'toolbar');
    this.toolbar.setAttribute('aria-label', 'Spreadsheet tools');
    this.toolbarItems = document.createElement('div');
    this.toolbarItems.className = 'os-toolbar-items';
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
    this.moreMenu.addEventListener('click', (event) => {
      if ((event.target as Element).closest('button')) this.setMoreOpen(false);
    });
    this.moreButton.setAttribute('aria-controls', this.moreMenu.id);
    more.append(this.moreButton, this.moreMenu);
    this.toolbar.append(this.toolbarItems, more);
    this.pluginToolbar = document.createElement('div');
    this.pluginToolbar.className = 'os-plugin-toolbar';
    this.namebox = document.createElement('input');
    this.namebox.className = 'os-namebox';
    this.namebox.setAttribute('aria-label', 'Cell address');
    this.namebox.value = 'A1';
    this.formula = document.createElement('input');
    this.formula.className = 'os-formula-input';
    this.formula.setAttribute('aria-label', 'Formula bar');
    this.formula.placeholder = 'Enter a value or formula, e.g. =SUM(D2:D10)';
    const bar = document.createElement('div');
    bar.className = 'os-formula-bar';
    const fx = document.createElement('span');
    fx.textContent = 'ƒx';
    bar.append(this.namebox, fx, this.formula);
    this.viewport = document.createElement('div');
    this.viewport.className = 'os-viewport';
    this.tabs = document.createElement('div');
    this.tabs.className = 'os-tabs';
    this.tabs.setAttribute('role', 'tablist');
    this.tabs.setAttribute('aria-label', 'Sheets');
    this.status = document.createElement('div');
    this.status.className = 'os-status';
    this.status.textContent = 'Ready';
    this.toast = document.createElement('div');
    this.toast.className = 'os-toast';
    this.toast.hidden = true;
    this.toast.setAttribute('role', 'status');
    const footer = document.createElement('div');
    footer.className = 'os-footer';
    footer.append(this.tabs, this.status);
    this.element.append(this.toolbar, bar, this.viewport, footer, this.toast);
    container.append(this.element);
    this.buildToolbar();
    this.toolbarItems.append(this.pluginToolbar);
    this.toolbar.hidden = options.toolbar === false;
    const resize = new ResizeObserver(() => this.layoutToolbar());
    resize.observe(this.toolbar);
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
    requestAnimationFrame(() => this.layoutToolbar());
    this.namebox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter')
        this.run(() => {
          this.selection.set({
            ...parseRange(this.namebox.value),
            sheetId: this.grid!.activeSheetId,
          });
          this.grid!.focus();
        });
    });
    this.formula.addEventListener('keydown', (e) => {
      if (e.key === 'Enter')
        this.run(() => {
          const s = this.selection.get()!;
          this.getWorkbook()
            .getSheetById(s.sheetId)!
            .range(address(s.startRow, s.startColumn))
            .setInput(this.formula.value);
          this.grid!.focus();
        });
      if (e.key === 'Escape') {
        this.updateSelection();
        this.grid?.focus();
      }
    });
    this.plugins = new PluginRegistry({
      getWorkbook: () => this.getWorkbook(),
      getSelection: () => this.selection.get(),
      addToolbar: (action) => {
        if (this.pluginToolbar.querySelector(`[data-action="${CSS.escape(action.id)}"]`))
          throw new OpenSheetError('PLUGIN_CONFLICT', 'Duplicate toolbar action');
        const b = this.button(action.label, action.run);
        b.dataset.action = action.id;
        this.pluginToolbar.append(b);
        this.layoutToolbar();
        return () => {
          b.remove();
          this.layoutToolbar();
        };
      },
      notify: (m) => this.notify(m),
      onError: (e) => {
        this.emit('plugin:error', e);
        this.error(e);
      },
    });
  }
  private active() {
    if (this.disposed) throw new OpenSheetError('DISPOSED', 'OpenSheet has been disposed');
  }
  ready() {
    this.active();
    return Promise.resolve();
  }
  createWorkbook(options: Parameters<typeof createSnapshot>[0] = {}) {
    this.active();
    if (this.book)
      throw new OpenSheetError('INVALID_ARGUMENT', 'Use load() to replace an existing workbook');
    this.attach(new Workbook(createSnapshot(options)));
    return this.book!;
  }
  async load(snapshot: WorkbookSnapshot) {
    this.active();
    const next = new Workbook(validateSnapshot(snapshot));
    const defs = this.pluginList.filter((p) => this.plugins.has(p.id));
    this.plugins.dispose();
    this.pluginList = [];
    this.grid?.dispose();
    this.stop?.();
    this.book?.dispose();
    this.attach(next);
    if (defs.length) this.use(defs);
  }
  getWorkbook() {
    this.active();
    if (!this.book) throw new OpenSheetError('INVALID_ARGUMENT', 'Create or load a workbook first');
    return this.book;
  }
  private attach(book: Workbook) {
    this.book = book;
    const read =
      this.options.mode === 'read' || book.toJSON().extensions['opensheet.readOnly'] === true;
    book.setReadOnly(read);
    this.formula.readOnly = read;
    this.grid = new CanvasGrid(this.viewport, book, {
      readOnly: read,
      onSelection: (s) => {
        this.updateSelection();
        this.emit('selection:changed', s);
      },
      onError: (e) => this.error(e),
    });
    this.stop = book.onCommit((c) => {
      this.renderTabs();
      this.updateSelection();
      this.emit('workbook:committed', c);
    });
    this.renderTabs();
    this.updateSelection();
  }
  setMode(mode: 'edit' | 'read') {
    this.active();
    this.options.mode = mode;
    const enforced = this.getWorkbook().toJSON().extensions['opensheet.readOnly'] === true;
    const read = mode === 'read' || enforced;
    this.book!.setReadOnly(read);
    this.grid?.setReadOnly(read);
    this.formula.readOnly = read;
  }
  selection = {
    get: (): Selection | null => {
      this.active();
      return this.grid?.selection ?? null;
    },
    set: (s: Selection) => {
      this.active();
      if (!this.grid) throw new OpenSheetError('INVALID_ARGUMENT', 'No workbook');
      this.grid.setSelection(s);
    },
    onChange: (fn: (s: Selection) => void) => this.on('selection:changed', fn),
  };
  commands = {
    execute: (command: Command | CommandEnvelope) => this.getWorkbook().execute(command),
    describe: () => Object.keys(commandSchemas),
    canExecute: (command: Command) => {
      try {
        if (
          this.options.mode === 'read' ||
          this.getWorkbook().toJSON().extensions['opensheet.readOnly'] === true
        )
          return false;
        return (
          commandSchemas[command.type as keyof typeof commandSchemas]?.safeParse(command.payload)
            .success ?? false
        );
      } catch {
        return false;
      }
    },
  };
  use(plugin: Plugin | Plugin[]) {
    this.active();
    const list = Array.isArray(plugin) ? plugin : [plugin];
    this.plugins.use(list);
    this.pluginList.push(...list);
    return this;
  }
  on<K extends keyof EventMap>(event: K, listener: (value: EventMap[K]) => void) {
    this.active();
    const set = this.events.get(event) ?? new Set();
    set.add(listener);
    this.events.set(event, set);
    return () => {
      set.delete(listener);
    };
  }
  private emit<K extends keyof EventMap>(event: K, value: EventMap[K]) {
    for (const fn of this.events.get(event) ?? []) {
      try {
        fn(value);
      } catch (e) {
        this.error(e);
      }
    }
  }
  notify(message: string) {
    this.toast.textContent = message;
    this.toast.hidden = false;
    clearTimeout(this.timer);
    this.timer = setTimeout(() => {
      this.toast.hidden = true;
    }, 5000);
  }
  private error(e: unknown) {
    this.notify(e instanceof Error ? e.message : String(e));
    this.options.onError?.(e);
  }
  private run(fn: () => unknown) {
    try {
      fn();
    } catch (e) {
      this.error(e);
    }
  }
  private button(label: string, run: () => void, title = label, icon?: IconNode) {
    const b = document.createElement('button');
    b.type = 'button';
    if (icon) {
      const svg = createElement(icon, {
        width: 16,
        height: 16,
        'aria-hidden': 'true',
        focusable: 'false',
      });
      b.append(svg);
      if (label) b.append(document.createTextNode(label));
    } else b.textContent = label;
    b.title = title;
    b.setAttribute('aria-label', title);
    b.addEventListener('click', () => this.run(run));
    return b;
  }
  private setMoreOpen(open: boolean) {
    this.moreMenu.hidden = !open;
    this.moreButton.setAttribute('aria-expanded', String(open));
  }
  private layoutToolbar() {
    if (this.disposed || this.toolbar.hidden || !this.toolbar.isConnected) return;
    this.setMoreOpen(false);
    this.toolbarItems.append(...this.moreMenu.children);
    this.moreButton.hidden = true;
    if (this.toolbarItems.scrollWidth <= this.toolbarItems.clientWidth) return;
    this.moreButton.hidden = false;
    while (this.toolbarItems.scrollWidth > this.toolbarItems.clientWidth) {
      const last = this.toolbarItems.lastElementChild;
      if (!last) break;
      this.moreMenu.prepend(last);
    }
  }
  private buildToolbar() {
    const group = (...buttons: HTMLElement[]) => {
      const g = document.createElement('div');
      g.className = 'os-tool-group';
      g.append(...buttons);
      this.toolbarItems.append(g);
    };
    const range = () => {
      const s = this.selection.get();
      if (!s) throw new Error('No selection');
      return this.getWorkbook().getSheetById(s.sheetId)!.range(s);
    };
    const style = (s: CellStyle) => range().setStyle(s);
    group(
      this.button('', () => this.getWorkbook().undo(), 'Undo', Undo2),
      this.button('', () => this.getWorkbook().redo(), 'Redo', Redo2),
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
    fmt.addEventListener('change', () => this.run(() => style({ numberFormat: fmt.value })));
    group(fmt);
    group(
      this.button(
        '',
        () => {
          const s = this.selection.get()!,
            c = this.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ bold: !this.getWorkbook().getStyle(c?.styleId).bold });
        },
        'Bold',
        Bold,
      ),
      this.button(
        '',
        () => {
          const s = this.selection.get()!,
            c = this.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ italic: !this.getWorkbook().getStyle(c?.styleId).italic });
        },
        'Italic',
        Italic,
      ),
      this.button(
        '',
        () => {
          const s = this.selection.get()!,
            c = this.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ underline: !this.getWorkbook().getStyle(c?.styleId).underline });
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
    color.addEventListener('input', () => this.run(() => style({ background: color.value })));
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
          const s = this.selection.get()!,
            c = this.getWorkbook().getCell(s.sheetId, s.startRow, s.startColumn);
          style({ wrap: !this.getWorkbook().getStyle(c?.styleId).wrap });
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
          const s = this.selection.get()!;
          this.getWorkbook().getSheetById(s.sheetId)!.insertRows(s.startRow);
        },
        'Insert row',
        Rows3,
      ),
      this.button(
        'Column',
        () => {
          const s = this.selection.get()!;
          this.getWorkbook().getSheetById(s.sheetId)!.insertColumns(s.startColumn);
        },
        'Insert column',
        Columns3,
      ),
    );
    group(
      this.button(
        'Freeze',
        () => {
          const s = this.selection.get()!,
            book = this.getWorkbook(),
            current = book.sheetData(s.sheetId).freeze;
          book.getSheetById(s.sheetId)!.setFreeze(current.rows ? 0 : 1, 0);
        },
        'Freeze first row',
        PanelTop,
      ),
    );
  }
  private renderTabs() {
    this.tabs.replaceChildren();
    for (const sheet of this.getWorkbook().getSheets()) {
      if (this.book!.sheetData(sheet.id).hidden) continue;
      const b = this.button(sheet.name, () => {
        this.grid!.setSheet(sheet.id);
        this.renderTabs();
      });
      b.setAttribute('role', 'tab');
      b.setAttribute('aria-selected', String(sheet.id === this.grid!.activeSheetId));
      this.tabs.append(b);
    }
    this.tabs.append(
      this.button(
        '',
        () => {
          const names = new Set(
            this.getWorkbook()
              .getSheets()
              .map((s) => s.name),
          );
          let i = 1;
          while (names.has(`Sheet${i}`)) i++;
          const sh = this.getWorkbook().addSheet(`Sheet${i}`);
          this.grid!.setSheet(sh.id);
          this.renderTabs();
        },
        'Add sheet',
        Plus,
      ),
    );
  }
  private updateSelection() {
    const s = this.grid?.selection;
    if (!s || !this.book) return;
    this.namebox.value = rangeAddress(s);
    const c = this.book.getCell(s.sheetId, s.startRow, s.startColumn);
    if (document.activeElement !== this.formula)
      this.formula.value =
        c?.input.type === 'formula'
          ? '=' + c.input.expression
          : c?.input.type === 'error'
            ? c.input.code
            : c && 'value' in c.input
              ? String(c.input.value)
              : '';
    const count = (s.endRow - s.startRow) * (s.endColumn - s.startColumn);
    let sum = 0,
      n = 0;
    if (count <= 10000) {
      for (const row of this.book.getSheetById(s.sheetId)!.range(s).getValues())
        for (const v of row)
          if (typeof v === 'number') {
            sum += v;
            n++;
          }
    }
    this.status.textContent = `${count > 1 ? `${count} cells${n ? ` · Sum ${sum.toLocaleString()}` : ''}` : 'Ready'}  ·  Local only`;
  }
  getGrid() {
    this.active();
    return this.grid;
  }
  dispose() {
    if (this.disposed) return;
    this.plugins.dispose();
    this.grid?.dispose();
    this.stop?.();
    this.book?.dispose();
    clearTimeout(this.timer);
    this.emit('lifecycle:disposed', undefined);
    this.events.clear();
    this.cleanups.forEach((fn) => fn());
    this.element.remove();
    this.disposed = true;
  }
}
export function createOpenSheet(options: OpenSheetOptions) {
  return new OpenSheet(options);
}
