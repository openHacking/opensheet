import type { OpenSheet } from 'opensheet';
import { createBindings } from './bindings.js';
import { $ } from './dom.js';

const sample = (sheet: string, range: string) =>
  `<button type="button" class="text-button feature-range" data-sheet="${sheet}" data-range="${range}">${sheet}!${range}</button>`;

export const featureGuide = `
  <details open>
    <summary>Edit cells and formulas</summary>
    <p>Change the discount in ${sample('Rates', 'B2')} and watch the order total below it recalculate.</p>
    <p>Select ${sample('Orders', 'E2')} to inspect its formula. Edit it in the formula bar for function help and highlighted references. Use Undo and Redo to replay changes.</p>
  </details>
  <details>
    <summary>Sort, filter, find and replace</summary>
    <p>Select ${sample('Orders', 'A2:G7')} and choose <b>Custom sort</b> in Sheet actions. Sort by column B (Region), then column C (Units).</p>
    <p>Select ${sample('Orders', 'F1:F7')} and choose <b>Custom filter</b>. Keep one header row, choose Equals and enter “Shipped”. Use Clear filter to show all orders again.</p>
    <p>Search for “Notebook” in Find. Use the arrows to move between matches, or choose Replace literal text in Sheet actions.</p>
  </details>
  <details>
    <summary>Copy, paste and fill</summary>
    <p>Copy ${sample('Practice', 'B2')}, then select ${sample('Practice', 'C2')} and try Paste or Paste special in Sheet actions. Compare formulas, values and formats.</p>
    <p>Select ${sample('Practice', 'A2:A6')} and choose Fill series in Sheet actions. Then select ${sample('Practice', 'B2:B6')} and use <b>Fill</b> on the toolbar to copy the formula down.</p>
  </details>
  <details>
    <summary>Format, merge and annotate</summary>
    <p>Try currency, percent or date in the Number format menu at ${sample('Practice', 'D2')}. Try borders, alignment, text styles, background and wrapping at ${sample('Practice', 'A9')}.</p>
    <p>Select ${sample('Practice', 'A8:C8')} to unmerge and remerge a heading. Select ${sample('Orders', 'A2')} and choose <b>Cell note and link</b> to edit its annotation.</p>
  </details>
  <details>
    <summary>Work with rows, columns and sheets</summary>
    <p>Scroll Orders to see its frozen headers. Select ${sample('Orders', 'C3')} to change the frozen area, insert, delete, resize, hide or show rows and columns from Sheet actions.</p>
    <p>Use the + next to the sheet tabs to add a sheet. Rename or delete the active sheet from Sheet actions. Try the Cell address field, Zoom, Read only, Table view and the keyboard shortcuts button.</p>
  </details>
  <details>
    <summary>Import, save and export</summary>
    <p>Save a native workbook, export JSON or export XLSX using the buttons above the editor. Reopen a downloaded file with Import file. XLSX compatibility warnings appear before a conversion with omitted features.</p>
    <p>Open <b>Code export</b> for LaTeX, Markdown, HTML, CSV and TSV. Try “Selected cells only”, then copy or download the result. Import file also accepts CSV and TSV.</p>
    <p>Reload to continue editing the sample in this browser. Use Reset demo to restore the template.</p>
  </details>
  <details>
    <summary>Try the plugin and other demos</summary>
    <p>Select ${sample('Orders', 'C2:C7')} and run <b>Σ Sum</b> on the toolbar. See the <a href="#plugins">plugin source</a> and <a href="#api">API quick start</a>.</p>
    <p>Try the <a href="#budget">budget simulator</a>, <a href="#sales">live sales chart</a>, <a href="#planner">project timeline</a>, <a href="#code">table-to-code demo</a> or <a href="#performance">Performance Lab</a>.</p>
  </details>`;

export function bindFeatureGuide(app: OpenSheet, notice: (message: string) => void) {
  const bindings = createBindings();
  bindings.on($('feature-guide'), 'click', (event) => {
    const button = (event.target as HTMLElement).closest<HTMLButtonElement>('button[data-range]');
    if (!button) return;
    try {
      const sheet = app.getWorkbook().getSheetByName(button.dataset.sheet!);
      if (!sheet) {
        notice('Sample sheet was renamed or removed. Reset demo to restore the template.');
        return;
      }
      app.selection.set({ sheetId: sheet.id, ...sheet.range(button.dataset.range!).bounds });
      document.querySelector<HTMLElement>('#spreadsheet .os-grid')?.focus();
    } catch (error) {
      notice(error instanceof Error ? error.message : String(error));
    }
  });
  return bindings.dispose;
}
