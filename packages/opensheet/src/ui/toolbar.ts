import { OpenSheetError, type CellStyle, type Selection, type Workbook } from '@opensheetjs/core';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowDown,
  Bold,
  Ellipsis,
  ArrowDownUp,
  Filter,
  StickyNote,
  Italic,
  Merge,
  PanelTop,
  Redo2,
  Split,
  Square,
  Underline,
  Undo2,
  WrapText,
  type IconNode,
} from 'lucide';
import type { ToolbarMenu } from './workbook-actions.js';
import { button } from './button.js';
let toolbarMenuId = 0;
export interface ToolbarActions {
  getWorkbook(): Workbook;
  getSelection(): Selection | null;
  getSelectionState(): { owner: object; version: number } | undefined;
  run(action: () => unknown): void;
  openMenu(kind: ToolbarMenu, trigger: HTMLElement): void;
  annotate(): void;
  openSelectionMenu(kind: 'cell' | 'row' | 'column', trigger: HTMLElement): void;
}
export class Toolbar {
  readonly element: HTMLDivElement;
  private items: HTMLDivElement;
  private moreButton: HTMLButtonElement;
  private moreMenu: HTMLDivElement;
  private dataTools: HTMLButtonElement[] = [];
  private filterButton?: HTMLButtonElement;
  private contextTools = document.createElement('div');
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
      false,
    );
    this.moreButton.setAttribute('aria-haspopup', 'true');
    this.moreButton.setAttribute('aria-expanded', 'false');
    this.moreButton.hidden = true;
    this.moreMenu = document.createElement('div');
    this.moreMenu.className = 'os-toolbar-menu';
    this.moreMenu.id = `os-toolbar-menu-${++toolbarMenuId}`;
    this.moreMenu.setAttribute('aria-label', 'More tools');
    this.contextTools.className = 'os-tool-group os-context-tools';
    for (const [kind, title] of [
      ['cell', 'Cell actions'],
      ['row', 'Row actions'],
      ['column', 'Column actions'],
    ] as const) {
      const tool = this.button(
        title,
        () => this.actions.openSelectionMenu(kind, this.moreButton),
        title,
        undefined,
        false,
      );
      tool.setAttribute('aria-haspopup', 'menu');
      this.contextTools.append(tool);
    }
    this.moreMenu.append(this.contextTools);
    this.moreMenu.hidden = true;
    this.moreMenu.addEventListener(
      'click',
      (event) => {
        if ((event.target as Element).closest('button')) {
          this.setMoreOpen(false);
          this.moreButton.focus({ preventScroll: true });
        }
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
    try {
      const resize = new ResizeObserver(() => this.layoutToolbar());
      this.cleanups.push(() => resize.disconnect());
      resize.observe(this.element);
      const dismissMore = (event: PointerEvent) => {
        if (!more.contains(event.target as Node)) this.setMoreOpen(false);
      };
      const escapeMore = (event: KeyboardEvent) => {
        if (event.key === 'Escape' && !this.moreMenu.hidden) {
          this.setMoreOpen(false);
          this.moreButton.focus({ preventScroll: true });
        }
      };
      document.addEventListener('pointerdown', dismissMore);
      this.cleanups.push(() => document.removeEventListener('pointerdown', dismissMore));
      document.addEventListener('keydown', escapeMore);
      this.cleanups.push(() => document.removeEventListener('keydown', escapeMore));
      this.frame = requestAnimationFrame(() => {
        this.frame = 0;
        this.layoutToolbar();
      });
    } catch (error) {
      try {
        this.dispose();
      } catch (cleanupError) {
        console.error(cleanupError);
      }
      throw error;
    }
  }
  private button(label: string, action: () => void, title = label, icon?: IconNode, stable = true) {
    return button(
      label,
      () => (stable ? this.runStable(action) : this.actions.run(action)),
      title,
      icon,
      this.abort.signal,
    );
  }
  private runStable(action: () => unknown) {
    const book = this.actions.getWorkbook();
    const revision = book.revision;
    const pending = book.persistenceState.pending;
    const selected = this.actions.getSelection();
    const selection = selected && { ...selected };
    const state = this.actions.getSelectionState();
    this.actions.run(() => {
      const current = this.actions.getSelection();
      const currentState = this.actions.getSelectionState();
      if (pending || book.persistenceState.pending)
        throw new OpenSheetError('BUSY', 'Wait for saving to finish, then retry the action.');
      if (
        this.actions.getWorkbook() !== book ||
        book.revision !== revision ||
        state?.owner !== currentState?.owner ||
        state?.version !== currentState?.version ||
        (selection === null) !== (current === null) ||
        (selection &&
          current &&
          (selection.sheetId !== current.sheetId ||
            selection.startRow !== current.startRow ||
            selection.endRow !== current.endRow ||
            selection.startColumn !== current.startColumn ||
            selection.endColumn !== current.endColumn))
      )
        throw new OpenSheetError(
          'REVISION_CONFLICT',
          'Selection or workbook changed. Select the range again.',
        );
      return action();
    });
  }
  private setMoreOpen(open: boolean) {
    this.moreMenu.hidden = !open;
    this.moreButton.setAttribute('aria-expanded', String(open));
  }
  private layoutToolbar() {
    if (this.disposed || this.element.hidden || !this.element.isConnected) return;
    this.setMoreOpen(false);
    this.items.append(
      ...[...this.moreMenu.children].filter((child) => child !== this.contextTools),
    );
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
      const book = this.actions.getWorkbook();
      if (book.persistenceState.pending)
        throw new OpenSheetError('BUSY', 'Wait for saving to finish, then retry the action.');
      return book.getSheetById(s.sheetId)!.range(s);
    };
    const style = async (s: CellStyle) => await range().setStyle(s);
    const toggle = async (property: 'bold' | 'italic' | 'underline' | 'wrap') => {
      const selected = this.actions.getSelection();
      if (!selected) throw new Error('No selection');
      const selection = { ...selected };
      const book = this.actions.getWorkbook();
      const revision = book.revision;
      const state = this.actions.getSelectionState();
      const cell = book.isLoaded(selection.sheetId, selection.startRow, selection.startColumn)
        ? book.peekCell(selection.sheetId, selection.startRow, selection.startColumn)
        : await book.getCell(selection.sheetId, selection.startRow, selection.startColumn);
      const currentState = this.actions.getSelectionState();
      if (
        this.disposed ||
        this.actions.getWorkbook() !== book ||
        book.revision !== revision ||
        state?.owner !== currentState?.owner ||
        state?.version !== currentState?.version
      )
        throw new OpenSheetError(
          'REVISION_CONFLICT',
          'Selection or workbook changed. Select the range again.',
        );
      const next = !book.getStyle(cell?.styleId)[property];
      await book.execute({
        protocolVersion: 1,
        commandId: crypto.randomUUID(),
        workbookId: book.id,
        baseRevision: revision,
        type: 'core.cells.style',
        payload: { sheetId: selection.sheetId, range: selection, style: { [property]: next } },
      });
    };
    group(
      this.button('', async () => await this.actions.getWorkbook().undo(), 'Undo', Undo2, false),
      this.button('', async () => await this.actions.getWorkbook().redo(), 'Redo', Redo2, false),
    );
    const menuButton = (kind: ToolbarMenu, label: string, icon: IconNode) => {
      const tool = this.button(
        '',
        () => this.actions.openMenu(kind, tool.getClientRects().length ? tool : this.moreButton),
        label,
        icon,
        false,
      );
      tool.dataset.menuLabel = 'true';
      tool.setAttribute('aria-haspopup', 'menu');
      tool.setAttribute('aria-expanded', 'false');
      return tool;
    };
    this.filterButton = menuButton('filter', 'Filter', Filter);
    const annotationButton = this.button(
      '',
      () => this.actions.annotate(),
      'Note & link',
      StickyNote,
      false,
    );
    annotationButton.dataset.menuLabel = 'true';
    this.dataTools = [
      menuButton('sort', 'Sort', ArrowDownUp),
      this.filterButton,
      menuButton('freeze', 'Freeze', PanelTop),
    ];
    this.updateDataTools(false, false);
    group(this.dataTools[0], this.dataTools[1], annotationButton, this.dataTools[2]);
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
    fmt.addEventListener('change', () => this.runStable(() => style({ numberFormat: fmt.value })), {
      signal: this.abort.signal,
    });
    group(fmt);
    group(
      this.button('', () => toggle('bold'), 'Bold', Bold),
      this.button('', () => toggle('italic'), 'Italic', Italic),
      this.button('', () => toggle('underline'), 'Underline', Underline),
    );
    const color = document.createElement('input');
    color.type = 'color';
    color.value = '#e1efe7';
    color.title = 'Cell background';
    color.setAttribute('aria-label', 'Cell background');
    color.addEventListener(
      'input',
      () => this.runStable(() => style({ background: color.value })),
      { signal: this.abort.signal },
    );
    group(
      color,
      this.button('', () => style({ align: 'left' }), 'Align left', AlignLeft),
      this.button('', () => style({ align: 'center' }), 'Align center', AlignCenter),
      this.button('', () => style({ align: 'right' }), 'Align right', AlignRight),
    );
    group(
      this.button('Merge', async () => await range().merge(), 'Merge', Merge),
      this.button('Unmerge', async () => await range().unmerge(), 'Unmerge', Split),
      this.button('Borders', () => style({ border: true }), 'Borders', Square),
      this.button('Wrap', () => toggle('wrap'), 'Wrap text', WrapText),
      this.button('Fill', async () => await range().fillDown(), 'Fill down', ArrowDown),
    );
  }
  updateDataTools(canEdit: boolean, filtered: boolean) {
    this.dataTools.forEach((tool) => (tool.disabled = !canEdit));
    this.filterButton?.classList.toggle('os-filter-active', filtered);
    this.filterButton?.setAttribute('aria-label', filtered ? 'Filter (active)' : 'Filter');
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
    let failure: unknown;
    let failed = false;
    for (const cleanup of this.cleanups) {
      try {
        cleanup();
      } catch (error) {
        if (!failed) failure = error;
        failed = true;
      }
    }
    if (failed) throw failure;
  }
}
