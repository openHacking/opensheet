import { address, validateSnapshot, type WorkbookSnapshot } from '@opensheetjs/core';
import { exportCell } from './cells.js';
import { add, finish, report } from './report.js';
import type { AdapterOptions, CompatibilityReport, SheetJSWorkbook } from './types.js';
export function toSheetJS(
  input: WorkbookSnapshot,
  options: AdapterOptions = {},
): { workbook: SheetJSWorkbook; report: CompatibilityReport } {
  const snapshot = validateSnapshot(input),
    r = report();
  const workbook: SheetJSWorkbook = {
    SheetNames: [],
    Sheets: Object.create(null),
    Workbook: {
      WBProps: { date1904: snapshot.dateSystem === '1904' },
      Sheets: [],
      CalcPr: { fullCalcOnLoad: true },
    },
  };
  for (const id of snapshot.sheetOrder) {
    const sh = snapshot.sheets.find((s) => s.id === id)!;
    const ws: Record<string, any> = {};
    let maxRow = 0,
      maxCol = 0;
    const rows = new Map(sh.rowOrder.map((id, i) => [id, i])),
      cols = new Map(sh.columnOrder.map((id, i) => [id, i]));
    for (const c of Object.values(sh.cells)) {
      const row = rows.get(c.rowId)!,
        col = cols.get(c.columnId)!;
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
      const x = exportCell(c, snapshot.styles, r, id, row, col);
      ws[address(row, col)] = x;
      r.summary.exact++;
    }
    for (const m of sh.merges) {
      maxRow = Math.max(maxRow, m.endRow - 1);
      maxCol = Math.max(maxCol, m.endColumn - 1);
    }
    ws['!ref'] = `A1:${address(maxRow, maxCol)}`;
    ws['!merges'] = sh.merges.map((m) => ({
      s: { r: m.startRow, c: m.startColumn },
      e: { r: m.endRow - 1, c: m.endColumn - 1 },
    }));
    ws['!rows'] = sh.rowOrder.map((id) => {
      const m = sh.rows[id];
      return m ? { hpx: m.size, hidden: m.hidden } : undefined;
    });
    ws['!cols'] = sh.columnOrder.map((id) => {
      const m = sh.columns[id];
      return m ? { wpx: m.size, hidden: m.hidden } : undefined;
    });
    workbook.SheetNames.push(sh.name);
    workbook.Sheets[sh.name] = ws;
    workbook.Workbook.Sheets.push({ name: sh.name, Hidden: sh.hidden ? 1 : 0 });
    if (sh.filter || sh.freeze.rows || sh.freeze.columns)
      add(r, {
        code: 'VIEW_NOT_EXPORTED',
        feature: 'view',
        message: 'Filters and frozen panes are preserved only in OpenSheet JSON',
        action: 'dropped',
        severity: 'warning',
        sheetId: id,
      });
  }
  const prior = snapshot.extensions['opensheet.importReport'] as CompatibilityReport | undefined;
  if (prior?.issues)
    for (const issue of prior.issues) if (issue.action !== 'preserved') add(r, issue);
  finish(r, options);
  return { workbook, report: r };
}
