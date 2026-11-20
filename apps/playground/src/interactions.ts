import { address, type OpenSheet } from 'opensheet';
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
  bindings.on($('replace-text'), 'click', () => {
    const selection = app.selection.get();
    const grid = app.getGrid();
    if (!selection || !grid) return;
    const book = app.getWorkbook();
    showReplaceDialog(book, selection.sheetId, findInput.value, {
      notice,
      modal,
      isCurrent: () => {
        try {
          return app.getWorkbook() === book && app.getGrid() === grid;
        } catch {
          return false;
        }
      },
    });
  });
  bindings.on($('filter-search'), 'click', async () => {
    const selection = app.selection.get();
    if (!selection) return;
    try {
      const book = app.getWorkbook();
      if (book.persistenceState.pending) {
        notice('Wait for saving to finish, then retry the filter.');
        return;
      }
      await book.getSheetById(selection.sheetId)!.setFilter({
        column: selection.startColumn,
        query: findInput.value,
      });
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
      'Enter / F2     Edit selected cell\nTab / Shift+Tab     Next / previous cell\nArrow keys     Move selection\nShift + arrows     Extend selection\nShift+F10 / Menu     Open context menu\nCtrl/Cmd + C / V     Copy / paste\nCtrl/Cmd + Z     Undo\nCtrl/Cmd + Shift + Z     Redo\nDelete / Backspace     Clear values\nEscape     Cancel editing\n\nPaste values starting with = as formulas. Use an apostrophe to enter literal formula text.\n\n† after a value indicates an imported formula cache that has not been recalculated.',
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
