import {
  createSheet,
  createSnapshot,
  key,
  LIMITS,
  OpenSheetError,
  parseAddress,
  validateSnapshot,
  type WorkbookSnapshot,
} from '@opensheetjs/core';
import { importCell } from './cells.js';
import { add, finish, report } from './report.js';
import type { AdapterOptions, CompatibilityReport, Issue, SheetJSWorkbook } from './types.js';
export function fromSheetJS(
  input: SheetJSWorkbook,
  options: AdapterOptions = {},
): { snapshot: WorkbookSnapshot; report: CompatibilityReport } {
  if (
    !input ||
    !Array.isArray(input.SheetNames) ||
    input.SheetNames.length < 1 ||
    input.SheetNames.length > LIMITS.sheets
  )
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Invalid number of sheets');
  const snapshot = createSnapshot();
  snapshot.sheets = [];
  snapshot.sheetOrder = [];
  snapshot.dateSystem = input.Workbook?.WBProps?.date1904 ? '1904' : '1900';
  const r = report();
  let total = 0;
  for (const [index, name] of input.SheetNames.entries()) {
    const ws = input.Sheets[name];
    if (!ws || typeof ws !== 'object')
      throw new OpenSheetError('INVALID_ARGUMENT', 'Missing worksheet');
    const entries: Array<[number, number, any]> = [];
    let maxRow = 1,
      maxCol = 1;
    const collect = (row: number, col: number, c: any) => {
      if (c === undefined || c === null) return;
      if (row >= LIMITS.rows || col >= LIMITS.columns)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Workbook exceeds preview limits');
      entries.push([row, col, c]);
      maxRow = Math.max(maxRow, row + 1);
      maxCol = Math.max(maxCol, col + 1);
    };
    const dense = Array.isArray(ws) ? ws : ws['!data'];
    if (dense) {
      if (dense.length > LIMITS.rows)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Dense worksheet too large');
      for (let row = 0; row < dense.length; row++) {
        if (!dense[row]) continue;
        if (dense[row].length > LIMITS.columns)
          throw new OpenSheetError('LIMIT_EXCEEDED', 'Too many columns');
        for (let col = 0; col < dense[row].length; col++) collect(row, col, dense[row][col]);
      }
    } else
      for (const [a, c] of Object.entries(ws)) {
        if (a.startsWith('!')) continue;
        const p = parseAddress(a);
        collect(p.row, p.column, c);
      }
    const merges = Array.isArray(ws['!merges']) ? ws['!merges'] : [];
    if (merges.length > 10000) throw new OpenSheetError('LIMIT_EXCEEDED', 'Too many merged ranges');
    for (const m of merges) {
      if (!m.s || !m.e) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid merge');
      maxRow = Math.max(maxRow, m.e.r + 1);
      maxCol = Math.max(maxCol, m.e.c + 1);
    }
    const rowMeta = ws['!rows'] ?? [],
      colMeta = ws['!cols'] ?? [];
    maxRow = Math.max(maxRow, rowMeta.length);
    maxCol = Math.max(maxCol, colMeta.length);
    const sh = createSheet(name, Math.max(100, maxRow), Math.max(26, maxCol));
    sh.hidden = !!input.Workbook?.Sheets?.[index]?.Hidden;
    snapshot.sheets.push(sh);
    snapshot.sheetOrder.push(sh.id);
    const issue = (
      code: string,
      feature: string,
      message: string,
      action: Issue['action'] = 'dropped',
      a?: string,
    ) =>
      add(r, {
        code,
        feature,
        message,
        action,
        severity: action === 'blocked' ? 'error' : 'warning',
        sheetId: sh.id,
        address: a,
      });
    for (const [row, col, c] of entries) {
      const { record, readOnly } = importCell(c, row, col, sh, snapshot.dateSystem, issue);
      if (readOnly) snapshot.extensions['opensheet.readOnly'] = true;
      sh.cells[key(record.rowId, record.columnId)] = record;
      r.summary.exact++;
    }
    sh.merges = merges.map((m: any) => ({
      startRow: m.s.r,
      startColumn: m.s.c,
      endRow: m.e.r + 1,
      endColumn: m.e.c + 1,
    }));
    rowMeta.forEach((m: any, i: number) => {
      if (m)
        sh.rows[sh.rowOrder[i]] = {
          ...(m.hpx ? { size: m.hpx } : m.hpt ? { size: (m.hpt * 96) / 72 } : {}),
          ...(m.hidden ? { hidden: true } : {}),
        };
    });
    colMeta.forEach((m: any, i: number) => {
      if (m) {
        sh.columns[sh.columnOrder[i]] = {
          ...(m.wpx ? { size: m.wpx } : m.wch ? { size: m.wch * 7 + 5 } : {}),
          ...(m.hidden ? { hidden: true } : {}),
        };
        if (m.wch && !m.wpx)
          issue(
            'WIDTH_APPROXIMATED',
            'column-width',
            'Character width approximated using 7 px glyphs',
            'approximated',
          );
      }
    });
    for (const k of Object.keys(ws))
      if (k.startsWith('!') && !['!ref', '!data', '!merges', '!rows', '!cols'].includes(k))
        issue('UNSUPPORTED_METADATA', k, `Worksheet metadata ${k} is not preserved`);
  }
  if (input.vbaraw)
    add(r, {
      code: 'MACROS_REMOVED',
      feature: 'macros',
      message: 'Macros are never executed or exported',
      action: 'dropped',
      severity: 'warning',
    });
  if (input.Workbook?.Names?.length)
    add(r, {
      code: 'NAMES_UNSUPPORTED',
      feature: 'defined-names',
      message: 'Defined names are not preserved',
      action: 'dropped',
      severity: 'warning',
    });
  add(r, {
    code: 'DATA_INTERCHANGE',
    feature: 'xlsx-container',
    message:
      'Data conversion only: embedded images, charts, macros, and unknown file parts are not preserved',
    action: 'preserved',
    severity: 'info',
  });
  snapshot.extensions['opensheet.importReport'] = JSON.parse(JSON.stringify(r));
  finish(r, options);
  return { snapshot: validateSnapshot(snapshot), report: r };
}
