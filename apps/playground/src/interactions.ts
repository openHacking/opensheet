import { showPasteDialog } from './paste-dialog.js';
import { showSeriesDialog } from './series-dialog.js';
import { showAxisSizeDialog } from './axis-size-dialog.js';
import { showAnnotationDialog } from './annotation-dialog.js';
import { showFilterDialog } from './filter-dialog.js';
import { address, type OpenSheet } from 'opensheet';
import { showSortDialog } from './sort-dialog.js';
import { showReplaceDialog } from './replace-dialog.js';
import { createBindings } from './bindings.js';
import { $ } from './dom.js';
import type { Feedback } from './feedback.js';
import type { PlaygroundState } from './state.js';
export function createInteractions(
  app: OpenSheet,
  state: PlaygroundState,
  { notice, modal, textModal }: Pick<Feedback, 'notice' | 'modal' | 'textModal'>,
) {
  const bindings = createBindings();
  bindings.on($<HTMLInputElement>('read-only'), 'change', (e) =>
    app.setMode((e.target as HTMLInputElement).checked ? 'read' : 'edit'),
  );
  bindings.on($<HTMLSelectElement>('zoom'), 'change', (e) =>
    app.getGrid()?.setZoom(Number((e.target as HTMLSelectElement).value)),
  );
  let searchEpoch = 0;
  let activeSearch: AbortController | undefined;
  let searchOrigin: ReturnType<OpenSheet['selection']['get']> = null;
  const findInput = $<HTMLInputElement>('find');
  const searchRegion = findInput.parentElement!;
  const pendingMessage = 'Searching… Press Escape to cancel.';
  let searchTimer: ReturnType<typeof setTimeout> | undefined;
  const clearSearch = () => {
    clearTimeout(searchTimer);
    searchTimer = undefined;
    searchRegion.removeAttribute('aria-busy');
    const status = $('notice');
    if (status.textContent === pendingMessage) status.hidden = true;
    activeSearch = undefined;
    searchOrigin = null;
  };
  const cancelStaleSearch = () => {
    if (!activeSearch) return;
    searchEpoch++;
    activeSearch.abort();
    clearSearch();
  };
  const stopSelection = app.selection.onChange((selection) => {
    if (
      activeSearch &&
      searchOrigin &&
      (selection.sheetId !== searchOrigin.sheetId ||
        selection.startRow !== searchOrigin.startRow ||
        selection.endRow !== searchOrigin.endRow ||
        selection.startColumn !== searchOrigin.startColumn ||
        selection.endColumn !== searchOrigin.endColumn)
    )
      cancelStaleSearch();
  });
  const stopCommit = app.on('workbook:committed', cancelStaleSearch);
  let previousSearch:
    | {
        book: ReturnType<OpenSheet['getWorkbook']>;
        revision: number;
        sheet: string;
        query: string;
        row: number;
        column: number;
      }
    | undefined;
  async function find(direction: 'next' | 'previous') {
    const query = findInput.value;
    cancelStaleSearch();
    if (!query) return;
    const request = ++searchEpoch;
    const controller = new AbortController();
    activeSearch = controller;
    searchRegion.setAttribute('aria-busy', 'true');
    searchTimer = setTimeout(() => {
      if (activeSearch === controller) notice(pendingMessage);
    }, 150);
    let current = () => request === searchEpoch;
    try {
      const selection = { ...app.selection.get()! },
        book = app.getWorkbook();
      searchOrigin = selection;
      const revision = book.revision;
      const continuation =
        previousSearch &&
        previousSearch.book === book &&
        previousSearch.revision === revision &&
        previousSearch.sheet === selection.sheetId &&
        previousSearch.query === query &&
        previousSearch.row === selection.startRow &&
        previousSearch.column === selection.startColumn &&
        selection.endRow === selection.startRow + 1 &&
        selection.endColumn === selection.startColumn + 1;
      current = () => {
        try {
          const now = app.selection.get();
          return (
            request === searchEpoch &&
            app.getWorkbook() === book &&
            book.revision === revision &&
            $<HTMLInputElement>('find').value === query &&
            !!now &&
            now.sheetId === selection.sheetId &&
            now.startRow === selection.startRow &&
            now.endRow === selection.endRow &&
            now.startColumn === selection.startColumn &&
            now.endColumn === selection.endColumn
          );
        } catch {
          return false;
        }
      };
      const result = await book.search(
        selection.sheetId,
        query,
        continuation ? state.foundIndex : -1,
        { signal: controller.signal, direction },
      );
      if (!current()) return;
      if (!result.match) {
        previousSearch = undefined;
        notice('No matching cells');
        return;
      }
      const { row, column } = result.match;
      state.foundIndex = row * book.sheetData(selection.sheetId).columnOrder.length + column;
      previousSearch = { book, revision, sheet: selection.sheetId, query, row, column };
      clearSearch();
      app.selection.set({
        sheetId: selection.sheetId,
        startRow: row,
        endRow: row + 1,
        startColumn: column,
        endColumn: column + 1,
      });
      notice(`${result.count} matching cells`);
    } catch (error) {
      if (current()) notice(String(error));
    } finally {
      if (activeSearch === controller) clearSearch();
    }
  }
  const startFromButton = (direction: 'next' | 'previous') => {
    findInput.focus({ preventScroll: true });
    void find(direction);
  };
  bindings.on($('find-next'), 'click', () => startFromButton('next'));
  bindings.on($('find-previous'), 'click', () => startFromButton('previous'));
  bindings.on(findInput, 'keydown', (e) => {
    if (e.isComposing || e.keyCode === 229) return;
    if (e.key === 'Escape' && activeSearch) {
      e.preventDefault();
      cancelStaleSearch();
      notice('Search cancelled');
      return;
    }
    if (e.key !== 'Enter') return;
    e.preventDefault();
    void find(e.shiftKey ? 'previous' : 'next');
  });
  bindings.on(findInput, 'input', () => {
    cancelStaleSearch();
    previousSearch = undefined;
    state.foundIndex = -1;
  });
  bindings.on($<HTMLSelectElement>('sheet-action'), 'change', async (e) => {
    const select = e.target as HTMLSelectElement,
      value = select.value;
    select.value = '';
    try {
      const book = app.getWorkbook(),
        s = app.selection.get()!,
        sh = book.getSheetById(s.sheetId)!;
      // These actions interpret coordinates or existing pane state from the current
      // view. A queued structural write can make that view stale before execution.
      if (
        book.persistenceState.pending &&
        [
          'sort-asc',
          'sort-desc',
          'freeze-selection',
          'freeze-first-row',
          'freeze-first-column',
          'unfreeze',
          'insert-row',
          'insert-column',
          'delete-row',
          'delete-column',
          'hide-row',
          'hide-column',
          'show-selected-rows',
          'show-selected-columns',
          'filter',
        ].includes(value)
      ) {
        notice('Wait for saving to finish, then retry the sheet action.');
        return;
      }
      const dialogGrid = app.getGrid();
      const dialogFeedback = {
        notice,
        modal,
        isCurrent: () => {
          try {
            return app.getWorkbook() === book && app.getGrid() === dialogGrid;
          } catch {
            return false;
          }
        },
      };
      switch (value) {
        case 'sort-custom':
          showSortDialog(book, s, dialogFeedback);
          break;
        case 'sort-asc':
        case 'sort-desc':
          await sh.sort(s, s.startColumn, value === 'sort-asc' ? 'asc' : 'desc');
          break;
        case 'freeze-selection':
        case 'freeze-first-row':
        case 'freeze-first-column':
        case 'unfreeze': {
          const current = book.sheetData(sh.id).freeze;
          const rows =
            value === 'freeze-selection'
              ? s.startRow
              : value === 'freeze-first-row'
                ? 1
                : value === 'unfreeze'
                  ? 0
                  : current.rows;
          const columns =
            value === 'freeze-selection'
              ? s.startColumn
              : value === 'freeze-first-column'
                ? 1
                : value === 'unfreeze'
                  ? 0
                  : current.columns;
          if (rows !== current.rows || columns !== current.columns)
            await sh.setFreeze(rows, columns);
          break;
        }
        case 'resize-row':
          showAxisSizeDialog(book, s, 'row', dialogFeedback);
          break;
        case 'resize-column':
          showAxisSizeDialog(book, s, 'column', dialogFeedback);
          break;
        case 'insert-row':
          await sh.insertRows(s.startRow, s.endRow - s.startRow);
          break;
        case 'insert-column':
          await sh.insertColumns(s.startColumn, s.endColumn - s.startColumn);
          break;
        case 'delete-row':
          await sh.deleteRows(s.startRow, s.endRow - s.startRow);
          break;
        case 'delete-column':
          await sh.deleteColumns(s.startColumn, s.endColumn - s.startColumn);
          break;
        case 'hide-row':
          await sh.setRowsHidden(s.startRow, s.endRow - s.startRow, true);
          break;
        case 'hide-column':
          await sh.setColumnsHidden(s.startColumn, s.endColumn - s.startColumn, true);
          break;
        case 'show-selected-rows':
          await sh.setRowsHidden(s.startRow, s.endRow - s.startRow, false);
          break;
        case 'show-selected-columns':
          await sh.setColumnsHidden(s.startColumn, s.endColumn - s.startColumn, false);
          break;
        case 'show-rows':
          await sh.setAllRowsHidden(false);
          break;
        case 'show-columns':
          await sh.setAllColumnsHidden(false);
          break;
        case 'paste-special': {
          const grid = app.getGrid();
          const selectionVersion = grid?.selectionVersion;
          showPasteDialog(book, s, dialogFeedback, (selection) => {
            if (
              grid &&
              app.getGrid() === grid &&
              app.getWorkbook() === book &&
              grid.selectionVersion === selectionVersion
            )
              app.selection.set(selection);
          });
          break;
        }
        case 'fill-series':
          await showSeriesDialog(book, s, dialogFeedback);
          break;
        case 'annotation':
          await showAnnotationDialog(book, s, dialogFeedback);
          break;
        case 'filter-custom':
          showFilterDialog(book, s.sheetId, s.startColumn, dialogFeedback);
          break;
        case 'filter':
          await sh.setFilter({ column: s.startColumn, query: $<HTMLInputElement>('find').value });
          break;
        case 'clear-filter':
          await sh.setFilter(null);
          break;
        case 'rename': {
          const next = prompt('Sheet name', sh.name);
          if (next) await book.renameSheet(sh.id, next);
          break;
        }
        case 'remove':
          if (confirm(`Delete sheet “${sh.name}”? You can undo this.`))
            await book.removeSheet(sh.id);
          break;
        case 'replace': {
          const query = $<HTMLInputElement>('find').value;
          showReplaceDialog(book, sh.id, query, dialogFeedback);
          break;
        }
      }
    } catch (error) {
      notice(error instanceof Error ? error.message : String(error));
    }
  });
  bindings.on($('accessible'), 'click', async () => {
    const s = app.selection.get()!,
      sheet = app.getWorkbook().getSheetById(s.sheetId)!,
      range = sheet.getUsedRange(),
      box = document.createElement('div'),
      table = document.createElement('table'),
      caption = document.createElement('caption');
    caption.textContent = `${sheet.name} — first 200 rows (read-only accessible view)`;
    table.append(caption);
    for (const row of await sheet
      .range({
        ...range.bounds,
        endRow: Math.min(200, range.bounds.endRow),
        endColumn: Math.min(80, range.bounds.endColumn),
      })
      .getDisplayValues()) {
      const tr = document.createElement('tr');
      for (const v of row) {
        const td = document.createElement('td');
        td.textContent = v;
        tr.append(td);
      }
      table.append(tr);
    }
    box.append(table);
    modal('Accessible table view', box);
  });
  bindings.on($('help'), 'click', () =>
    textModal(
      'Keyboard shortcuts',
      'Enter / F2     Edit selected cell\nTab / Shift+Tab     Next / previous cell\nArrow keys     Move selection\nShift + arrows     Extend selection\nCtrl/Cmd + C / V     Copy / paste\nCtrl/Cmd + Z     Undo\nCtrl/Cmd + Shift + Z     Redo\nDelete / Backspace     Clear values\nEscape     Cancel editing\n\nPaste values starting with = as formulas. Use an apostrophe to enter literal formula text.\n\n† after a value indicates an imported formula cache that has not been recalculated.',
    ),
  );
  return {
    dispose() {
      cancelStaleSearch();
      stopSelection();
      stopCommit();
      bindings.dispose();
    },
  };
}
