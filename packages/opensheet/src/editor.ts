import {
  OpenSheetError,
  Workbook,
  commandSchemas,
  createWorkbookFile,
  openWorkbook,
  uid,
  type WorkbookOptions,
  type Command,
  type CommandEnvelope,
  type Selection,
  type WorkbookFile,
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
import { WorkbookActions } from './ui/workbook-actions.js';
export class OpenSheet {
  readonly element: HTMLDivElement;
  readonly plugins: PluginRegistry;
  private book?: Workbook;
  private grid?: CanvasGrid;
  private toolbarView: Toolbar;
  private workbookActions: WorkbookActions;
  private formulaBar: FormulaBar;
  private tabsView: SheetTabs;
  private statusBar = new StatusBar();
  private toastView = new Toast();
  private viewport: HTMLDivElement;
  private stop?: () => void;
  private disposed = false;
  private disposing = false;
  private pendingActions = 0;
  private formulaDraft = false;
  private cellDraft = false;
  private importing = false;
  private events = new EditorEvents((error) => this.error(error));
  constructor(private options: OpenSheetOptions) {
    const container =
      typeof options.container === 'string'
        ? document.querySelector<HTMLElement>(options.container)
        : options.container;
    if (!container) throw new OpenSheetError('INVALID_ARGUMENT', 'Container was not found');
    this.element = document.createElement('div');
    this.element.className = 'opensheet';
    const rollback: Array<() => void> = [
      () => this.element.remove(),
      () => this.events.clear(),
      () => this.toastView.dispose(),
    ];
    try {
      this.workbookActions = new WorkbookActions(this.element, {
        book: () => (this.disposed ? undefined : this.book),
        grid: () => (this.disposed ? undefined : this.grid),
        select: (selection) => this.selection.set(selection),
        canEdit: () => !this.importing && this.options.mode !== 'read',
        preserveFocus: () =>
          !!this.grid?.isComposing || this.hasDraft || this.formulaBar.canPickReference(),
        run: (fn) => this.run(fn),
        notice: (message) => this.notify(message),
      });
      rollback.push(() => this.workbookActions.dispose());
      this.toolbarView = new Toolbar(
        {
          getWorkbook: () => this.getWorkbook(),
          getSelection: () => this.selection.get(),
          getSelectionState: () =>
            this.grid ? { owner: this.grid, version: this.grid.selectionVersion } : undefined,
          run: (fn) => this.run(fn),
          openMenu: (kind, trigger) => this.workbookActions.toolbar(kind, trigger),
          annotate: () => this.workbookActions.annotate(),
          openSelectionMenu: (kind, trigger) => this.workbookActions.selectionMenu(kind, trigger),
        },
        options.toolbar !== false,
      );
      rollback.push(() => this.toolbarView.dispose());
      this.formulaBar = new FormulaBar({
        book: () => this.getWorkbook(),
        selection: () => this.selection.get(),
        select: (s) => this.selection.set(s),
        move: (rows, columns, extend) => this.grid?.moveSelection(rows, columns, extend),
        references: (selections) => this.grid?.setFormulaReferences(selections),
        draft: (dirty) => this.setDraft('formula', dirty),
        focus: () => this.grid?.focus(),
        run: (fn) => this.run(fn),
      });
      rollback.push(() => this.formulaBar.dispose());
      this.tabsView = new SheetTabs({
        book: () => this.getWorkbook(),
        activeSheet: () => this.grid!.activeSheetId,
        setSheet: (id) => this.grid!.setSheet(id),
        openMenu: (id, trigger) => this.workbookActions.sheetMenu(id, trigger),
        preserveFormulaFocus: () =>
          !!this.grid?.isComposing ||
          this.formulaBar.canPickReference() ||
          this.transferCellFormulaDraft(),
        run: (fn) => this.run(fn),
      });
      rollback.push(() => this.tabsView.dispose());
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
      rollback.push(() => this.plugins.dispose());
      container.append(this.element);
    } catch (error) {
      this.disposed = true;
      for (const cleanup of rollback.reverse()) {
        try {
          cleanup();
        } catch (cleanupError) {
          console.error(cleanupError);
        }
      }
      throw error;
    }
  }
  private active() {
    if (this.disposed) throw new OpenSheetError('DISPOSED', 'OpenSheet has been disposed');
  }
  get hasDraft() {
    return this.formulaDraft || this.cellDraft;
  }
  private setDraft(kind: 'formula' | 'cell', dirty: boolean) {
    const previous = this.hasDraft;
    if (kind === 'formula') this.formulaDraft = dirty;
    else this.cellDraft = dirty;
    if (previous !== this.hasDraft && !this.disposed && !this.disposing)
      this.emit('editor:draft', this.hasDraft);
  }
  private noPendingWrites() {
    if (this.hasDraft || this.book?.persistenceState.pending)
      throw new OpenSheetError('BUSY', 'Wait for pending changes before replacing the workbook');
  }
  private async abandonPendingReplacement(next: Workbook) {
    if (!this.hasDraft && !this.book?.persistenceState.pending) return;
    try {
      await next.deleteStorage();
    } finally {
      next.dispose();
    }
    throw new OpenSheetError('BUSY', 'Wait for pending changes before replacing the workbook');
  }
  ready() {
    this.active();
    return this.grid?.ready() ?? Promise.resolve();
  }
  async createWorkbook(
    options: Parameters<typeof createWorkbookFile>[0] = {},
    storage: WorkbookOptions = {},
  ) {
    this.active();
    if (this.book)
      throw new OpenSheetError('INVALID_ARGUMENT', 'Use load() to replace an existing workbook');
    const next = await new Workbook(createWorkbookFile(options), storage).ready();
    if (this.disposed) {
      try {
        await next.deleteStorage();
      } finally {
        next.dispose();
      }
      throw new OpenSheetError('DISPOSED', 'Editor closed during initialization');
    }
    this.attach(next);
    await this.grid!.ready();
    return this.book!;
  }
  async load(snapshot: WorkbookFile, options: { signal?: AbortSignal } = {}) {
    this.active();
    this.noPendingWrites();
    options.signal?.throwIfAborted();
    const previous = this.book;
    const revision = previous?.revision;
    const next = await new Workbook({ ...snapshot, workbookId: uid('wb') }).ready();
    if (options.signal?.aborted || this.book !== previous || previous?.revision !== revision) {
      try {
        await next.deleteStorage();
      } finally {
        next.dispose();
      }
      options.signal?.throwIfAborted();
      throw new OpenSheetError('BUSY', 'Workbook changed during import');
    }
    if (this.disposed) {
      try {
        await next.deleteStorage();
      } finally {
        next.dispose();
      }
      throw new OpenSheetError('DISPOSED', 'Editor closed during import');
    }
    if (this.hasDraft || this.book?.persistenceState.pending)
      await this.abandonPendingReplacement(next);
    const defs = this.plugins.definitions();
    this.plugins.dispose();
    this.grid?.dispose();
    this.stop?.();
    this.book?.dispose();
    this.attach(next);
    this.restorePlugins(defs);
    await this.grid!.ready();
  }
  private async importWorkbook(action: (book: Workbook) => Promise<void>, signal?: AbortSignal) {
    this.active();
    this.noPendingWrites();
    if (this.importing) throw new OpenSheetError('BUSY', 'Another import is running');
    const book = this.getWorkbook();
    this.importing = true;
    this.grid?.setReadOnly(true);
    this.formulaBar.setReadOnly(true);
    this.updateDataTools();
    try {
      await action(book);
      signal?.throwIfAborted();
      this.active();
      if (this.book !== book) throw new OpenSheetError('BUSY', 'Workbook changed during import');
      await this.attachWorkbook(book);
    } finally {
      this.importing = false;
      if (!this.disposed && this.book) {
        this.grid?.setReadOnly(this.book.isReadOnly);
        this.formulaBar.setReadOnly(this.book.isReadOnly);
        this.updateDataTools();
      }
    }
  }
  async loadJSON(
    input: WorkbookFile | string | Blob | ReadableStream<Uint8Array>,
    options: { signal?: AbortSignal } = {},
  ) {
    await this.importWorkbook((book) => book.importJSON(input, options), options.signal);
  }
  async loadBinary(
    input: Blob | Uint8Array | ReadableStream<Uint8Array>,
    options: { signal?: AbortSignal } = {},
  ) {
    await this.importWorkbook((book) => book.importBinary(input, options), options.signal);
  }
  async open(id: string, storage: WorkbookOptions = {}) {
    this.active();
    this.noPendingWrites();
    const initial = this.book;
    if (this.book?.id === id) {
      this.grid?.dispose();
      this.stop?.();
      await this.book.close();
      if (this.book !== initial)
        throw new OpenSheetError('BUSY', 'Workbook changed during restore');
      this.book = undefined;
    }
    const previous = this.book;
    const revision = previous?.revision;
    const next = await openWorkbook(id, storage);
    if (this.disposed) {
      next.dispose();
      throw new OpenSheetError('DISPOSED', 'Editor closed during restore');
    }
    if (
      this.book !== previous ||
      previous?.revision !== revision ||
      this.hasDraft ||
      this.book?.persistenceState.pending
    ) {
      next.dispose();
      throw new OpenSheetError('BUSY', 'Workbook changed or has pending edits during restore');
    }
    const defs = this.plugins.definitions();
    this.plugins.dispose();
    this.grid?.dispose();
    this.stop?.();
    this.book?.dispose();
    this.attach(next);
    this.restorePlugins(defs);
    await this.grid!.ready();
    return next;
  }
  async attachWorkbook(book: Workbook) {
    this.active();
    this.noPendingWrites();
    const previous = this.book;
    const revision = previous?.revision;
    await book.ready();
    this.active();
    this.noPendingWrites();
    if (this.book !== previous || previous?.revision !== revision)
      throw new OpenSheetError('BUSY', 'Workbook changed during attachment');
    // A target can be closed while its asynchronous readiness is pending.
    // Validate it before dismantling the current editor and workbook.
    void book.metadata;
    const defs = this.plugins.definitions();
    this.plugins.dispose();
    this.grid?.dispose();
    this.stop?.();
    if (this.book !== book) this.book?.dispose();
    this.attach(book);
    this.restorePlugins(defs);
    await this.grid!.ready();
  }
  getWorkbook() {
    this.active();
    if (!this.book) throw new OpenSheetError('INVALID_ARGUMENT', 'Create or load a workbook first');
    return this.book;
  }
  private attach(book: Workbook) {
    const zoom = this.grid?.getZoom() ?? 1;
    this.book = book;
    const read =
      this.options.mode === 'read' || book.metadata.extensions['opensheet.readOnly'] === true;
    book.setReadOnly(read);
    const readonly = this.importing || book.isReadOnly;
    this.formulaBar.setReadOnly(readonly);
    this.workbookActions.dismissMenu();
    this.grid = new CanvasGrid(this.viewport, book, {
      onContextMenu:
        this.options.contextMenu === false
          ? undefined
          : (target, point) => this.workbookActions.gridMenu(target, point),
      preserveContextFocus: () => this.hasDraft || this.formulaBar.canPickReference(),
      readOnly: readonly,
      onKeyboardReference: () =>
        this.transferCellFormulaDraft() && this.formulaBar.startKeyboardReference(),
      referencePicker: {
        begin: () => this.formulaBar.beginReference(),
        update: (selection) => this.formulaBar.updateReference(selection),
        end: () => this.formulaBar.endReference(),
      },
      onSelection: (s) => {
        this.updateSelection();
        this.emit('selection:changed', s);
      },
      onError: (e) => this.error(e),
      onDraft: (dirty) => this.setDraft('cell', dirty),
    });
    if (zoom !== 1) this.grid.setZoom(zoom);
    const stopCommit = book.onCommit((c) => {
      this.tabsView.render();
      this.updateSelection();
      this.emit('workbook:committed', c);
    });
    const stopData = book.onData(() => this.updateSelection());
    let viewReadOnly = readonly;
    const stopPersistence = book.onPersistence((state) => {
      const nextReadOnly = this.importing || book.isReadOnly;
      this.updateDataTools();
      if (nextReadOnly !== viewReadOnly) {
        viewReadOnly = nextReadOnly;
        this.grid?.setReadOnly(nextReadOnly);
        this.formulaBar.setReadOnly(nextReadOnly);
      }
      this.emit('workbook:persistence', state);
    });
    this.stop = () => {
      stopCommit();
      stopData();
      stopPersistence();
    };
    this.emit('workbook:persistence', book.persistenceState);
    this.tabsView.render();
    this.updateSelection();
  }
  setMode(mode: 'edit' | 'read') {
    this.active();
    this.options.mode = mode;
    if (!this.book) return;
    const enforced = this.getWorkbook().metadata.extensions['opensheet.readOnly'] === true;
    const read = mode === 'read' || enforced;
    this.book!.setReadOnly(read);
    this.grid?.setReadOnly(this.importing || this.book!.isReadOnly);
    this.formulaBar.setReadOnly(this.importing || this.book!.isReadOnly);
    this.updateSelection();
  }
  private transferCellFormulaDraft() {
    const draft = this.grid?.takeFormulaDraft();
    if (!draft) return false;
    this.selection.set(draft.origin);
    this.formulaBar.adoptDraft(draft.value, draft.start, draft.end, draft.identity, draft.revision);
    return true;
  }
  selection = {
    get: (): Selection | null => {
      this.active();
      return this.grid?.selection ?? null;
    },
    set: (s: Selection) => {
      this.active();
      if (!this.grid) throw new OpenSheetError('INVALID_ARGUMENT', 'No workbook');
      const previousSheet = this.grid.activeSheetId;
      this.grid.setSelection(s);
      if (previousSheet !== this.grid.activeSheetId) this.tabsView.render();
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
          this.getWorkbook().isReadOnly ||
          this.getWorkbook().metadata.extensions['opensheet.readOnly'] === true
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
    return this;
  }
  /** Workbook replacement has already happened; plugin failures must not misreport it as failed. */
  private restorePlugins(definitions: Plugin[]) {
    const pending = new Map(definitions.map((plugin) => [plugin.id, plugin]));
    while (pending.size) {
      const ready = [...pending.values()].find((plugin) =>
        (plugin.requires ?? []).every((id) => !pending.has(id)),
      );
      if (!ready) {
        this.reportPluginRestoreError(
          new OpenSheetError('PLUGIN_CONFLICT', 'Plugin dependency cycle during restoration'),
        );
        return;
      }
      pending.delete(ready.id);
      const missing = (ready.requires ?? []).find((id) => !this.plugins.has(id));
      if (missing) {
        this.reportPluginRestoreError(
          new OpenSheetError(
            'PLUGIN_CONFLICT',
            `Plugin ${ready.id} could not restore dependency ${missing}`,
          ),
        );
        continue;
      }
      try {
        this.use(ready);
      } catch {
        // PluginRegistry already reports setup failures and rolls back their resources.
      }
    }
  }
  private reportPluginRestoreError(error: unknown) {
    try {
      this.emit('plugin:error', error);
      this.error(error);
    } catch {
      // Observers cannot undo a completed workbook replacement.
    }
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
    // Pending actions may settle after disposal; their host and toast no longer exist.
    if (this.disposed) return;
    this.notify(e instanceof Error ? e.message : String(e));
    try {
      const result: unknown = this.options.onError?.(e);
      if (result && typeof (result as PromiseLike<unknown>).then === 'function')
        void Promise.resolve(result).catch((error) => console.error(error));
    } catch (error) {
      console.error(error);
    }
  }
  private run(fn: () => unknown) {
    this.pendingActions++;
    this.element.setAttribute('aria-busy', 'true');
    this.statusBar.element.textContent = 'Working…';
    void Promise.resolve()
      .then(fn)
      .catch((e) => this.error(e))
      .finally(() => {
        this.pendingActions--;
        if (!this.pendingActions && !this.disposed) {
          this.element.setAttribute('aria-busy', 'false');
          this.updateSelection();
        }
      });
  }
  private updateDataTools() {
    const book = this.book;
    const sheetId = this.grid?.activeSheetId;
    const editable =
      !!book && !!sheetId && !this.importing && !book.isReadOnly && !book.persistenceState.pending;
    const filtered =
      !!book && !!sheetId && !!book.getSheetById(sheetId) && !!book.sheetData(sheetId).filter;
    this.toolbarView.updateDataTools(editable, filtered);
  }
  private updateSelection() {
    if (!this.book) return;
    const s = this.grid?.selection ?? null;
    this.updateDataTools();
    this.formulaBar.update(this.book, s);
    void this.statusBar.update(this.book, s).catch((e) => {
      if (!this.disposed) this.error(e);
    });
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
    this.workbookActions.dispose();
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
