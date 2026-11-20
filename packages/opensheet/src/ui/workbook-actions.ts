import { OpenSheetError, type Selection, type Workbook } from '@opensheetjs/core';
import {
  pasteClipboard,
  type CanvasGrid,
  type GridTarget,
  type PasteOptions,
} from '@opensheetjs/renderer';
import { ActionMenu, type MenuItem } from './action-menu.js';
import { showAnnotationDialog } from './annotation-dialog.js';
import { showAxisSizeDialog } from './axis-size-dialog.js';
import { readClipboard } from './clipboard-access.js';
import { showFilterDialog } from './filter-dialog.js';
import { showPasteDialog } from './paste-dialog.js';
import { showSeriesDialog } from './series-dialog.js';
import { showSortDialog } from './sort-dialog.js';
import { WorkbookDialog } from './workbook-dialog.js';

export type ToolbarMenu = 'sort' | 'filter' | 'freeze';
type PasteMode = NonNullable<PasteOptions['mode']>;
type Axis = 'row' | 'column';

interface ActionHost {
  book(): Workbook | undefined;
  grid(): CanvasGrid | undefined;
  select(selection: Selection): void;
  canEdit(): boolean;
  preserveFocus(): boolean;
  run(action: () => unknown): void;
  notice(message: string): void;
}

export class WorkbookActions {
  private menu: ActionMenu;
  private dialog: WorkbookDialog;
  constructor(
    container: HTMLElement,
    private host: ActionHost,
  ) {
    this.menu = new ActionMenu(host.run);
    this.dialog = new WorkbookDialog(container);
  }

  private context() {
    const book = this.host.book();
    const grid = this.host.grid();
    if (!book || !grid) return;
    const selection = { ...grid.selection };
    const revision = book.revision;
    const version = grid.selectionVersion;
    const sheet = book.getSheetById(selection.sheetId)!;
    const attached = () => this.host.book() === book && this.host.grid() === grid;
    const current = () => attached() && grid.selectionVersion === version;
    const canWrite = (queue = false) =>
      this.host.canEdit() && !book.isReadOnly && (queue || !book.persistenceState.pending);
    const check = (write: boolean, queue = false) => {
      if (!current() || book.revision !== revision)
        throw new OpenSheetError(
          'REVISION_CONFLICT',
          'Selection or workbook changed. Select the range again.',
        );
      if (write && !canWrite(queue))
        throw new OpenSheetError(
          'BUSY',
          'Editing is unavailable. Check read only and saving status.',
        );
    };
    const item = (
      label: string,
      action: () => unknown,
      write = true,
      unavailable = false,
      queue = false,
    ): MenuItem => ({
      label,
      disabled: unavailable || (write && !canWrite(queue)),
      action: () => {
        check(write, queue);
        return action();
      },
    });
    const feedback = {
      modal: (title: string, form: HTMLElement) => this.dialog.show(title, form),
      notice: this.host.notice,
      isCurrent: attached,
      isTargetCurrent: current,
    };
    return { book, grid, selection, revision, sheet, current, canWrite, check, item, feedback };
  }

  private sort(
    ctx: NonNullable<ReturnType<WorkbookActions['context']>>,
    column: number,
    direction: 'asc' | 'desc',
    confirmRange: boolean,
    custom: boolean,
  ) {
    const { book, selection, sheet, feedback } = ctx;
    const range = confirmRange ? sheet.getUsedRange().bounds : selection;
    if (column < range.startColumn || column >= range.endColumn) {
      this.host.notice('Select a data range that includes this sort column.');
      return;
    }
    if (custom || confirmRange)
      return showSortDialog(book, { ...range, sheetId: sheet.id }, feedback, {
        column,
        direction,
        suggestRange: confirmRange,
      });
    return sheet.sort(range, column, direction);
  }

  private annotation(
    ctx: NonNullable<ReturnType<WorkbookActions['context']>>,
    target?: Extract<GridTarget, { kind: 'cell' }>,
  ) {
    const { book, selection, feedback } = ctx;
    const { row, column } = target ?? { row: selection.startRow, column: selection.startColumn };
    return showAnnotationDialog(
      book,
      {
        sheetId: selection.sheetId,
        startRow: row,
        endRow: row + 1,
        startColumn: column,
        endColumn: column + 1,
      },
      feedback,
    );
  }

  annotate() {
    const ctx = this.context();
    if (!ctx || this.host.preserveFocus()) return;
    this.host.run(ctx.item('Note & link', () => this.annotation(ctx), false).action!);
  }

  toolbar(kind: ToolbarMenu, trigger: HTMLElement) {
    const ctx = this.context();
    if (!ctx || this.host.preserveFocus()) return;
    const { selection, book, sheet, item, feedback } = ctx;
    const data = book.sheetData(sheet.id);
    let items: MenuItem[];
    if (kind === 'sort') {
      const axis = selection.endRow - selection.startRow < 2 || ctx.grid.selectionAxis === 'column';
      items = [
        item('Sort ascending', () => this.sort(ctx, selection.startColumn, 'asc', axis, false)),
        item('Sort descending', () => this.sort(ctx, selection.startColumn, 'desc', axis, false)),
        item('Custom sort…', () => this.sort(ctx, selection.startColumn, 'asc', axis, true)),
      ];
    } else if (kind === 'filter') {
      items = [
        item(data.filter ? 'Edit filter…' : 'Create filter…', () =>
          showFilterDialog(book, sheet.id, selection.startColumn, feedback),
        ),
        item('Clear filter', () => sheet.setFilter(null), true, !data.filter),
      ];
    } else {
      const { rows, columns } = data.freeze;
      const setFreeze = (nextRows: number, nextColumns: number) => {
        if (rows !== nextRows || columns !== nextColumns)
          return sheet.setFreeze(nextRows, nextColumns);
      };
      items = [
        item('Freeze first row', () => setFreeze(1, columns)),
        item('Freeze first column', () => setFreeze(rows, 1)),
        item('Freeze before selection', () => setFreeze(selection.startRow, selection.startColumn)),
        item('Unfreeze rows and columns', () => setFreeze(0, 0), true, !rows && !columns),
      ];
    }
    const rect = trigger.getBoundingClientRect();
    this.menu.open(
      kind === 'sort' ? 'Sort' : kind === 'filter' ? 'Filter' : 'Freeze',
      items,
      { x: rect.left, y: rect.bottom },
      () => ctx.grid.focus(),
      trigger,
    );
  }

  gridMenu(target: GridTarget, point: { x: number; y: number }, trigger?: HTMLElement) {
    const ctx = this.context();
    if (!ctx || this.host.preserveFocus()) return;
    const { book, grid, selection, item, sheet, feedback } = ctx;
    let items: MenuItem[];
    if (target.kind === 'cell') {
      const paste = (label: string, mode: PasteMode, keyboard = false) =>
        item(label, () => this.paste(ctx, mode, keyboard));
      items = [
        item('Copy', () => grid.copyToClipboard(), false),
        paste('Paste', 'all'),
        {
          label: 'Paste special',
          disabled: !ctx.canWrite(),
          items: [
            paste('Paste all', 'all'),
            paste('Values only', 'values'),
            paste('Formula inputs only', 'formulas'),
            paste('Formats only', 'formats'),
            paste('Paste from keyboard…', 'all', true),
          ],
        },
        item('Clear contents', () => sheet.range(selection).clear()),
        item('Fill series…', () => showSeriesDialog(book, selection, feedback)),
        item('Note & link…', () => this.annotation(ctx, target), false),
      ];
    } else {
      items = this.axisItems(ctx, target.kind, target.index);
    }
    const label =
      target.kind === 'cell'
        ? 'Cell actions'
        : target.kind === 'row'
          ? 'Row actions'
          : 'Column actions';
    this.menu.open(label, items, point, () => grid.focus(), trigger);
  }

  private axisItems(
    ctx: NonNullable<ReturnType<WorkbookActions['context']>>,
    kind: Axis,
    index: number,
  ): MenuItem[] {
    const { selection, sheet, item, book, feedback } = ctx;
    const row = kind === 'row';
    const start = row ? selection.startRow : selection.startColumn;
    const end = row ? selection.endRow : selection.endColumn;
    const count = end - start;
    const label = `${count} ${kind}${count === 1 ? '' : 's'}`;
    const insert = (index: number) =>
      row ? sheet.insertRows(index, count) : sheet.insertColumns(index, count);
    const resize = row ? 'Set selected row height…' : 'Set selected column width…';
    const visibility = (hidden: boolean) =>
      row
        ? sheet.setRowsHidden(start, count, hidden)
        : sheet.setColumnsHidden(start, count, hidden);
    const actions = [
      item(`Insert ${label} ${row ? 'above' : 'left'}`, () => insert(start)),
      item(`Insert ${label} ${row ? 'below' : 'right'}`, () => insert(end)),
      item(
        `Delete ${label}`,
        () => (row ? sheet.deleteRows(start, count) : sheet.deleteColumns(start, count)),
        true,
        count === (row ? sheet.rowCount : sheet.columnCount),
      ),
      item(resize, () => showAxisSizeDialog(book, selection, kind, feedback)),
      item(`Hide ${label}`, () => visibility(true)),
      item(`Show selected ${kind}s`, () => visibility(false)),
      item(
        `Show all ${kind}s`,
        () => (row ? sheet.setAllRowsHidden(false) : sheet.setAllColumnsHidden(false)),
        true,
        false,
        true,
      ),
    ];
    if (!row)
      actions.push(
        item('Sort ascending by this column', () => this.sort(ctx, index, 'asc', true, false)),
        item('Sort descending by this column', () => this.sort(ctx, index, 'desc', true, false)),
        item('Filter this column…', () => showFilterDialog(book, sheet.id, index, feedback)),
      );
    return actions;
  }

  selectionMenu(kind: 'cell' | Axis, trigger: HTMLElement) {
    const grid = this.host.grid();
    if (!grid || this.host.preserveFocus()) return;
    const { startRow, startColumn } = grid.selection;
    const target: GridTarget =
      kind === 'cell'
        ? { kind, row: startRow, column: startColumn }
        : { kind, index: kind === 'row' ? startRow : startColumn };
    this.gridMenu(target, grid.contextPoint(target), trigger);
  }

  sheetMenu(id: string, trigger: HTMLElement) {
    const ctx = this.context();
    const sheet = ctx?.book.getSheetById(id);
    if (!ctx || !sheet || this.host.preserveFocus()) return;
    const rect = trigger.getBoundingClientRect();
    this.menu.open(
      'Sheet tab actions',
      [
        ctx.item('Rename sheet…', async () => {
          const name = prompt('Sheet name', sheet.name);
          if (!name) return;
          ctx.check(true);
          await ctx.book.renameSheet(id, name);
        }),
        ctx.item(
          'Delete sheet…',
          async () => {
            if (!confirm(`Delete sheet “${sheet.name}”? You can undo this.`)) return;
            ctx.check(true);
            await ctx.book.removeSheet(id);
          },
          true,
          ctx.book.getSheets().filter((sheet) => !ctx.book.sheetData(sheet.id).hidden).length <= 1,
        ),
      ],
      { x: rect.left, y: rect.top },
      () => ctx.grid.focus(),
      trigger,
    );
  }

  private async paste(
    ctx: NonNullable<ReturnType<WorkbookActions['context']>>,
    mode: PasteMode,
    keyboard: boolean,
  ) {
    const { book, revision, selection, feedback, current } = ctx;
    const select = (result: Selection) => {
      if (current()) this.host.select(result);
    };
    const data = keyboard ? undefined : await readClipboard();
    ctx.check(true);
    if (!data) {
      showPasteDialog(book, selection, feedback, select, mode);
      return;
    }
    const result = await pasteClipboard(book, selection, data, { mode, baseRevision: revision });
    select(result);
    this.host.notice('Clipboard pasted');
  }

  dismissMenu() {
    this.menu.close(false);
  }

  dispose() {
    this.menu.dispose();
    this.dialog.element.remove();
  }
}
