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
  function findNext() {
    const query = $<HTMLInputElement>('find').value.toLowerCase();
    if (!query) return;
    const s = app.selection.get()!,
      sheet = app.getWorkbook().getSheetById(s.sheetId)!,
      r = sheet.getUsedRange().bounds,
      matches: Array<[number, number]> = [];
    for (let row = 0; row < r.endRow; row++)
      for (let c = 0; c < r.endColumn; c++)
        if (app.getWorkbook().display(s.sheetId, row, c).toLowerCase().includes(query))
          matches.push([row, c]);
    if (!matches.length) {
      notice('No matching cells');
      return;
    }
    state.foundIndex = (state.foundIndex + 1) % matches.length;
    const [row, col] = matches[state.foundIndex];
    app.selection.set({
      sheetId: s.sheetId,
      startRow: row,
      endRow: row + 1,
      startColumn: col,
      endColumn: col + 1,
    });
    notice(`${state.foundIndex + 1} of ${matches.length} matches`);
  }
  bindings.on($('find-next'), 'click', findNext);
  bindings.on($('find'), 'keydown', (e) => {
    if (e.key === 'Enter') findNext();
  });
  bindings.on($('find'), 'input', () => {
    state.foundIndex = -1;
  });
  bindings.on($<HTMLSelectElement>('sheet-action'), 'change', (e) => {
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
          sh.sort(s, s.startColumn, value === 'sort-asc' ? 'asc' : 'desc');
          break;
        case 'delete-row':
          sh.deleteRows(s.startRow, s.endRow - s.startRow);
          break;
        case 'delete-column':
          sh.deleteColumns(s.startColumn, s.endColumn - s.startColumn);
          break;
        case 'hide-row':
          sh.setRowHidden(s.startRow, true);
          break;
        case 'show-rows':
          book.transaction({}, () => {
            const data = book.sheetData(sh.id);
            data.rowOrder.forEach((id, i) => {
              if (data.rows[id]?.hidden) sh.setRowHidden(i, false);
            });
          });
          break;
        case 'filter':
          sh.setFilter({ column: s.startColumn, query: $<HTMLInputElement>('find').value });
          break;
        case 'clear-filter':
          sh.setFilter(null);
          break;
        case 'rename': {
          const next = prompt('Sheet name', sh.name);
          if (next) book.renameSheet(sh.id, next);
          break;
        }
        case 'remove':
          if (confirm(`Delete sheet “${sh.name}”? You can undo this.`)) book.removeSheet(sh.id);
          break;
        case 'replace': {
          const query = $<HTMLInputElement>('find').value;
          if (!query) {
            notice('Enter search text first');
            return;
          }
          const replacement = prompt(`Replace “${query}” with:`);
          if (replacement === null) return;
          book.transaction({}, () => {
            const data = book.sheetData(sh.id);
            for (const c of Object.values(data.cells))
              if (c.input.type === 'string' && c.input.value.includes(query)) {
                const row = data.rowOrder.indexOf(c.rowId),
                  col = data.columnOrder.indexOf(c.columnId);
                sh.range(address(row, col)).setValues([
                  [c.input.value.replaceAll(query, replacement)],
                ]);
              }
          });
          break;
        }
      }
    } catch (error) {
      notice(error instanceof Error ? error.message : String(error));
    }
  });
  bindings.on($('accessible'), 'click', () => {
    const s = app.selection.get()!,
      sheet = app.getWorkbook().getSheetById(s.sheetId)!,
      range = sheet.getUsedRange(),
      box = document.createElement('div'),
      table = document.createElement('table'),
      caption = document.createElement('caption');
    caption.textContent = `${sheet.name} — first 200 rows (read-only accessible view)`;
    table.append(caption);
    for (const row of range.getDisplayValues().slice(0, 200)) {
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
  const api = `import { createOpenSheet } from 'opensheet';\nimport 'opensheet/style.css';\n\nconst app = createOpenSheet({ container: '#sheet' });\nconst book = app.createWorkbook();\nconst sheet = book.getSheets()[0];\n\nsheet.range('A1:B2').setValues([\n  ['Product', 'Revenue'], ['OpenSheet', 1200]\n]);\nsheet.range('B3').setFormulas([['SUM(B2:B2)']]);\napp.on('workbook:committed', ({ revision }) => {\n  console.log(revision, book.toJSON());\n});\n\n// On unmount\napp.dispose();`;
  bindings.on($('nav-api'), 'click', () => textModal('API quick start', api));
  bindings.on($('nav-plugins'), 'click', () =>
    textModal(
      'The Σ Sum button is a plugin',
      `definePlugin({\n  id: 'example.selection-sum',\n  version: '0.1.0',\n  apiVersion: '^0.1.0',\n  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],\n  setup(ctx) {\n    ctx.ui.toolbar.add({\n      id: 'sum', label: 'Σ Sum',\n      run() {\n        const selection = ctx.selection.get();\n        const rows = ctx.workbook.readRange(selection).values;\n        const sum = rows.flat().reduce(\n          (total, n) => total + (typeof n === 'number' ? n : 0), 0\n        );\n        ctx.ui.notify(String(sum));\n      }\n    });\n  }\n});`,
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
