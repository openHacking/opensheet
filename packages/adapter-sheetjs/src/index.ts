import {
  createSnapshot,
  createSheet,
  validateSnapshot,
  parseAddress,
  address,
  key,
  LIMITS,
  scalarValue,
  type WorkbookSnapshot,
  type Scalar,
  type CellInput,
  type CellRecord,
  OpenSheetError,
} from '@opensheetjs/core';
export type SheetJSWorkbook = {
  SheetNames: string[];
  Sheets: Record<string, any>;
  Workbook?: any;
  vbaraw?: unknown;
};
export type Issue = {
  code: string;
  severity: 'info' | 'warning' | 'error';
  sheetId?: string;
  address?: string;
  feature: string;
  action: 'preserved' | 'approximated' | 'dropped' | 'blocked';
  message: string;
  count?: number;
};
export type CompatibilityReport = {
  adapterVersion: string;
  sourceFormat: string;
  summary: { exact: number; approximated: number; dropped: number; blocked: number };
  issues: Issue[];
};
export type AdapterOptions = { unsupported?: 'report' | 'strict' };
const errors: Record<number, string> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
};
function report(): CompatibilityReport {
  return {
    adapterVersion: '0.1.0',
    sourceFormat: 'SheetJS',
    summary: { exact: 0, approximated: 0, dropped: 0, blocked: 0 },
    issues: [],
  };
}
function add(r: CompatibilityReport, issue: Issue) {
  const field = issue.action === 'preserved' ? 'exact' : issue.action;
  r.summary[field]++;
  const prior = r.issues.find((x) => x.code === issue.code && x.sheetId === issue.sheetId);
  if (prior) prior.count = (prior.count ?? 1) + 1;
  else if (r.issues.length < 200) r.issues.push(issue);
}
function finish(r: CompatibilityReport, o: AdapterOptions) {
  if (o.unsupported === 'strict' && r.issues.some((i) => i.action !== 'preserved'))
    throw new OpenSheetError('UNSUPPORTED_FEATURE', 'Strict compatibility check failed', r);
}
function dateSerial(value: Date | string, system: '1900' | '1904'): number {
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid date');
  const local = Date.UTC(
    d.getFullYear(),
    d.getMonth(),
    d.getDate(),
    d.getHours(),
    d.getMinutes(),
    d.getSeconds(),
    d.getMilliseconds(),
  );
  if (system === '1904') return (local - Date.UTC(1904, 0, 1)) / 86400000;
  const days = (local - Date.UTC(1899, 11, 31)) / 86400000;
  return days + (days >= 60 ? 1 : 0);
}
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
      if (++total > LIMITS.cells || row >= LIMITS.rows || col >= LIMITS.columns)
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
      let value: Scalar;
      switch (c.t) {
        case 'n':
          if (c.v === undefined && c.f) {
            value = { type: 'blank' };
            break;
          }
          if (typeof c.v !== 'number' || !Number.isFinite(c.v))
            throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid numeric cell');
          value = { type: 'number', value: c.v };
          break;
        case 'b':
          value = { type: 'boolean', value: !!c.v };
          break;
        case 'e':
          value = { type: 'error', code: errors[c.v] ?? '#VALUE!' };
          break;
        case 'd':
          value = { type: 'number', value: dateSerial(c.v, snapshot.dateSystem) };
          issue(
            'DATE_NORMALIZED',
            'date',
            'Date converted to local spreadsheet serial',
            'approximated',
            address(row, col),
          );
          break;
        case 'z':
          value = { type: 'blank' };
          break;
        default:
          value = c.v === undefined ? { type: 'blank' } : { type: 'string', value: String(c.v) };
      }
      const record: CellRecord = {
        rowId: sh.rowOrder[row],
        columnId: sh.columnOrder[col],
        input: c.f ? { type: 'formula', expression: String(c.f) } : value,
      };
      if (c.f && c.v !== undefined) record.cached = value;
      if (c.z) record.numberFormat = String(c.z);
      else if (c.t === 'd') record.numberFormat = 'yyyy-mm-dd';
      if (c.l?.Target) {
        if (/^(https?:|mailto:)/i.test(c.l.Target))
          record.link = {
            target: c.l.Target,
            ...(c.l.Tooltip ? { tooltip: String(c.l.Tooltip) } : {}),
          };
        else issue('UNSAFE_LINK', 'link', 'Unsupported hyperlink scheme omitted');
      }
      if (Array.isArray(c.c)) record.note = c.c.map((x: any) => String(x.t ?? '')).join('\n');
      if (c.s && Object.keys(c.s).length)
        issue('STYLE_NOT_IMPORTED', 'style', 'Advanced Excel styling is not imported');
      if (c.r || c.h)
        issue(
          'RICH_TEXT_FLATTENED',
          'rich-text',
          'Rich text flattened to plain text',
          'approximated',
        );
      if (c.F || c.D) {
        issue(
          'ARRAY_FORMULA',
          'array-formula',
          'Array formula workbook is read-only in this release',
          'blocked',
        );
        snapshot.extensions['opensheet.readOnly'] = true;
      }
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
      const x: Record<string, any> = {};
      const v = c.input;
      if (v.type === 'formula') {
        x.f = v.expression;
        x.t = 'n';
        add(r, {
          code: 'RECALC_REQUIRED',
          feature: 'formula',
          message: 'Formula result requires recalculation in the target application',
          action: 'approximated',
          severity: 'warning',
          sheetId: id,
          address: address(row, col),
        });
      } else if (v.type === 'blank') x.t = 'z';
      else if (v.type === 'error') {
        x.t = 'e';
        x.v = Number(Object.entries(errors).find(([, name]) => name === v.code)?.[0] ?? 15);
      } else {
        x.t = v.type === 'number' ? 'n' : v.type === 'boolean' ? 'b' : 's';
        x.v = v.value;
      }
      if (c.numberFormat) x.z = c.numberFormat;
      if (c.link && /^(https?:|mailto:)/i.test(c.link.target))
        x.l = { Target: c.link.target, Tooltip: c.link.tooltip };
      if (c.note) x.c = [{ a: 'OpenSheet', t: c.note }];
      if (c.styleId) {
        const style = snapshot.styles[c.styleId];
        if (style.numberFormat) x.z = style.numberFormat;
        if (Object.keys(style).some((k) => k !== 'numberFormat'))
          add(r, {
            code: 'STYLE_NOT_EXPORTED',
            feature: 'style',
            message:
              'SheetJS CE export does not guarantee visual styles; use OpenSheet JSON to preserve styles',
            action: 'dropped',
            severity: 'warning',
            sheetId: id,
          });
      }
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
