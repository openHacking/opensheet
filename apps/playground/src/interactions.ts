import { address, type OpenSheet } from 'opensheet';
import { createBindings } from './bindings.js';
import { $ } from './dom.js';
import type { Feedback } from './feedback.js';
import type { PlaygroundState } from './state.js';
export function createInteractions(
  app: OpenSheet,
  state: PlaygroundState,
  { notice, modal, textModal }: Pick<Feedback, 'notice' | 'modal' | 'textModal'>,
  showPanel: (panel: 'code' | 'insights') => void,
) {
  const bindings = createBindings();
  bindings.on($<HTMLInputElement>('read-only'), 'change', (e) =>
    app.setMode((e.target as HTMLInputElement).checked ? 'read' : 'edit'),
  );
  bindings.on($<HTMLSelectElement>('zoom'), 'change', (e) =>
    app.getGrid()?.setZoom(Number((e.target as HTMLSelectElement).value)),
  );
  async function findNext() {
    const query = $<HTMLInputElement>('find').value;
    if (!query) return;
    try {
      const selection = app.selection.get()!,
        book = app.getWorkbook();
      const result = await book.search(selection.sheetId, query, state.foundIndex);
      if (!result.match) {
        notice('No matching cells');
        return;
      }
      const { row, column } = result.match;
      state.foundIndex = row * book.sheetData(selection.sheetId).columnOrder.length + column;
      app.selection.set({
        sheetId: selection.sheetId,
        startRow: row,
        endRow: row + 1,
        startColumn: column,
        endColumn: column + 1,
      });
      notice(`${result.count} matching cells`);
    } catch (error) {
      notice(String(error));
    }
  }
  bindings.on($('find-next'), 'click', findNext);
  bindings.on($('find'), 'keydown', (e) => {
    if (e.key === 'Enter') findNext();
  });
  bindings.on($('find'), 'input', () => {
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
      switch (value) {
        case 'sort-asc':
        case 'sort-desc':
          await sh.sort(s, s.startColumn, value === 'sort-asc' ? 'asc' : 'desc');
          break;
        case 'delete-row':
          await sh.deleteRows(s.startRow, s.endRow - s.startRow);
          break;
        case 'delete-column':
          await sh.deleteColumns(s.startColumn, s.endColumn - s.startColumn);
          break;
        case 'hide-row':
          await sh.setRowHidden(s.startRow, true);
          break;
        case 'show-rows':
          await book.transaction({}, async (book) => {
            const data = book.sheetData(sh.id);
            for (let i = 0; i < data.rowOrder.length; i++)
              if (data.rows[data.rowOrder[i]]?.hidden)
                await book.getSheetById(sh.id)!.setRowHidden(i, false);
          });
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
          if (!query) {
            notice('Enter search text first');
            return;
          }
          const replacement = prompt(`Replace “${query}” with:`);
          if (replacement === null) return;
          await book.replaceText(sh.id, query, replacement);
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
  const api = `import { createOpenSheet } from 'opensheet';\nimport 'opensheet/style.css';\n\nconst app = createOpenSheet({ container: '#sheet' });\nconst book = await app.createWorkbook();\nconst sheet = book.getSheets()[0];\n\nawait sheet.range('A1:B2').setValues([\n  ['Product', 'Revenue'], ['OpenSheet', 1200]\n]);\nawait sheet.range('B3').setFormulas([['SUM(B2:B2)']]);\napp.on('workbook:committed', ({ revision }) => {\n  console.log(revision);\n});\n\n// On unmount\napp.dispose();`;
  bindings.on($('nav-api'), 'click', () => textModal('API quick start', api));
  bindings.on($('nav-plugins'), 'click', () =>
    textModal(
      'The Σ Sum button is a plugin',
      `definePlugin({\n  id: 'example.selection-sum',\n  version: '0.1.0',\n  apiVersion: '^0.1.0',\n  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],\n  setup(ctx) {\n    ctx.ui.toolbar.add({\n      id: 'sum', label: 'Σ Sum',\n      async run() {\n        const selection = ctx.selection.get();\n        const rows = (await ctx.workbook.readRange(selection)).values;\n        const sum = rows.flat().reduce(\n          (total, n) => total + (typeof n === 'number' ? n : 0), 0\n        );\n        ctx.ui.notify(String(sum));\n      }\n    });\n  }\n});`,
    ),
  );
  bindings.on($('help'), 'click', () =>
    textModal(
      'Keyboard shortcuts',
      'Enter / F2     Edit selected cell\nTab / Shift+Tab     Next / previous cell\nArrow keys     Move selection\nShift + arrows     Extend selection\nCtrl/Cmd + C / V     Copy / paste\nCtrl/Cmd + Z     Undo\nCtrl/Cmd + Shift + Z     Redo\nDelete / Backspace     Clear values\nEscape     Cancel editing\n\nPaste values starting with = as formulas. Use an apostrophe to enter literal formula text.\n\n† after a value indicates an imported formula cache that has not been recalculated.',
    ),
  );
  bindings.on($('nav-code'), 'click', () => showPanel('code'));
  bindings.on($('nav-sheet'), 'click', () =>
    document.querySelector<HTMLElement>('#spreadsheet .os-grid')?.focus(),
  );
  bindings.on($('nav-templates'), 'click', () => $<HTMLSelectElement>('scene-select').focus());
  return { dispose: bindings.dispose };
}
