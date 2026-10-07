import {
  OpenSheetError,
  Workbook,
  commandSchemas,
  createSnapshot,
  validateSnapshot,
  type Command,
  type CommandEnvelope,
  type Selection,
  type WorkbookSnapshot,
} from '@opensheetjs/core';
import { PluginRegistry, type Plugin } from '@opensheetjs/plugin-sdk';
import { CanvasGrid } from '@opensheetjs/renderer';
import { EditorEvents } from './events.js';
import type { EventMap, OpenSheetOptions } from './types.js';
import { FormulaBar } from './ui/formula-bar.js';
import { StatusBar } from './ui/status.js';
import { SheetTabs } from './ui/tabs.js';
import { Toast } from './ui/toast.js';
import { Toolbar } from './ui/toolbar.js';
export class OpenSheet {
  readonly element: HTMLDivElement;
  readonly plugins: PluginRegistry;
  private book?: Workbook;
  private grid?: CanvasGrid;
  private toolbarView: Toolbar;
  private formulaBar: FormulaBar;
  private tabsView: SheetTabs;
  private statusBar = new StatusBar();
  private toastView = new Toast();
  private viewport: HTMLDivElement;
  private stop?: () => void;
  private disposed = false;
  private disposing = false;
  private pluginList: Plugin[] = [];
  private events = new EditorEvents((error) => this.error(error));
  constructor(private options: OpenSheetOptions) {
    const container =
      typeof options.container === 'string'
        ? document.querySelector<HTMLElement>(options.container)
        : options.container;
    if (!container) throw new OpenSheetError('INVALID_ARGUMENT', 'Container was not found');
    this.element = document.createElement('div');
    this.element.className = 'opensheet';
    this.toolbarView = new Toolbar(
      {
        getWorkbook: () => this.getWorkbook(),
        getSelection: () => this.selection.get(),
        run: (fn) => this.run(fn),
      },
      options.toolbar !== false,
    );
    this.formulaBar = new FormulaBar({
      book: () => this.getWorkbook(),
      selection: () => this.selection.get(),
      select: (s) => this.selection.set(s),
      focus: () => this.grid?.focus(),
      run: (fn) => this.run(fn),
    });
    this.tabsView = new SheetTabs({
      book: () => this.getWorkbook(),
      activeSheet: () => this.grid!.activeSheetId,
      setSheet: (id) => this.grid!.setSheet(id),
      run: (fn) => this.run(fn),
    });
    this.viewport = document.createElement('div');
    this.viewport.className = 'os-viewport';
    const footer = document.createElement('div');
    footer.className = 'os-footer';
    footer.append(this.tabsView.element, this.statusBar.element);
    this.element.append(
      this.toolbarView.element,
      this.formulaBar.element,
      this.viewport,
      footer,
      this.toastView.element,
    );
    container.append(this.element);
    this.plugins = new PluginRegistry({
      getWorkbook: () => this.getWorkbook(),
      getSelection: () => this.selection.get(),
      addToolbar: (action) => this.toolbarView.add(action),
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
    this.formulaBar.setReadOnly(read);
    this.grid = new CanvasGrid(this.viewport, book, {
      readOnly: read,
      onSelection: (s) => {
        this.updateSelection();
        this.emit('selection:changed', s);
      },
      onError: (e) => this.error(e),
    });
    this.stop = book.onCommit((c) => {
      this.tabsView.render();
      this.updateSelection();
      this.emit('workbook:committed', c);
    });
    this.tabsView.render();
    this.updateSelection();
  }
  setMode(mode: 'edit' | 'read') {
    this.active();
    this.options.mode = mode;
    const enforced = this.getWorkbook().toJSON().extensions['opensheet.readOnly'] === true;
    const read = mode === 'read' || enforced;
    this.book!.setReadOnly(read);
    this.grid?.setReadOnly(read);
    this.formulaBar.setReadOnly(read);
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
    return this.events.on(event, listener);
  }
  private emit<K extends keyof EventMap>(event: K, value: EventMap[K]) {
    this.events.emit(event, value);
  }
  notify(message: string) {
    this.toastView.show(message);
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
  private updateSelection() {
    if (!this.book) return;
    const s = this.grid?.selection ?? null;
    this.formulaBar.update(this.book, s);
    this.statusBar.update(this.book, s);
  }
  getGrid() {
    this.active();
    return this.grid;
  }
  dispose() {
    if (this.disposed || this.disposing) return;
    this.disposing = true;
    this.plugins.dispose();
    this.grid?.dispose();
    this.stop?.();
    this.book?.dispose();
    this.toastView.dispose();
    this.emit('lifecycle:disposed', undefined);
    this.events.clear();
    this.toolbarView.dispose();
    this.formulaBar.dispose();
    this.tabsView.dispose();
    this.element.remove();
    this.disposed = true;
  }
}
export function createOpenSheet(options: OpenSheetOptions) {
  return new OpenSheet(options);
}
