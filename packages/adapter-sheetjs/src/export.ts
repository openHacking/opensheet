import { address, FileReader, LIMITS, OpenSheetError, type WorkbookFile } from '@opensheetjs/core';
import {
  exportCell,
  hasXlsxControlText,
  hasXlsxEscapeText,
  hasXlsxUnpairedSurrogate,
  XLSX_CELL_TEXT_LIMIT,
} from './cells.js';
import { add, finish, report } from './report.js';
import { createPrecedenceNormalizer } from './formula.js';
import { unsupportedXlsxSheetName, XLSX_SHEET_NAME_MESSAGE } from './sheet-name.js';
import type { AdapterOptions, CompatibilityReport, SheetJSWorkbook } from './types.js';
export function toSheetJS(
  input: WorkbookFile,
  options: AdapterOptions = {},
): { workbook: SheetJSWorkbook; report: CompatibilityReport } {
  const { workbook, report } = buildSheetJS(input, options, false);
  return { workbook, report };
}

export function canWriteFreeze(
  freeze: { rows: number; columns: number },
  rows: number,
  columns: number,
): boolean {
  return (
    (freeze.rows === 0 || (freeze.rows < rows && freeze.rows < LIMITS.rows)) &&
    (freeze.columns === 0 || (freeze.columns < columns && freeze.columns < LIMITS.columns))
  );
}

function validatePriorReport(prior: CompatibilityReport | undefined): void {
  if (prior !== undefined) {
    if (
      !prior ||
      typeof prior !== 'object' ||
      !Array.isArray(prior.issues) ||
      prior.issues.length > 200 ||
      !prior.summary ||
      typeof prior.summary !== 'object'
    )
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid original import compatibility report');
    for (const field of ['exact', 'approximated', 'dropped', 'blocked'] as const)
      if (!Number.isSafeInteger(prior.summary[field]) || prior.summary[field] < 0)
        throw new OpenSheetError(
          'INVALID_ARGUMENT',
          'Invalid original import compatibility report',
        );
    const retained = { exact: 0, approximated: 0, dropped: 0, blocked: 0 };
    for (const issue of prior.issues) {
      if (
        !issue ||
        typeof issue !== 'object' ||
        typeof issue.code !== 'string' ||
        typeof issue.feature !== 'string' ||
        typeof issue.message !== 'string' ||
        !['info', 'warning', 'error'].includes(issue.severity) ||
        (issue.sheetId !== undefined && typeof issue.sheetId !== 'string') ||
        (issue.address !== undefined && typeof issue.address !== 'string') ||
        !['preserved', 'approximated', 'dropped', 'blocked'].includes(issue.action) ||
        (issue.count !== undefined && (!Number.isSafeInteger(issue.count) || issue.count < 1))
      )
        throw new OpenSheetError(
          'INVALID_ARGUMENT',
          'Invalid original import compatibility report',
        );
      const field = issue.action === 'preserved' ? 'exact' : issue.action;
      retained[field] += issue.count ?? 1;
      if (!Number.isSafeInteger(retained[field]) || retained[field] > prior.summary[field])
        throw new OpenSheetError(
          'INVALID_ARGUMENT',
          'Invalid original import compatibility report',
        );
    }
    if (
      prior.omittedIssueOccurrences !== undefined &&
      (!Number.isSafeInteger(prior.omittedIssueOccurrences) || prior.omittedIssueOccurrences < 0)
    )
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid original import compatibility report');
    if (prior.omittedIssueOccurrences !== undefined) {
      const omitted = BigInt(prior.omittedIssueOccurrences);
      const unlistedLosses = (['approximated', 'dropped', 'blocked'] as const).reduce(
        (total, action) => total + BigInt(prior.summary[action] - retained[action]),
        0n,
      );
      const availableOccurrences = unlistedLosses + BigInt(prior.summary.exact - retained.exact);
      if (omitted < unlistedLosses || omitted > availableOccurrences)
        throw new OpenSheetError(
          'INVALID_ARGUMENT',
          'Invalid original import compatibility report',
        );
    }
  }
}

/** Internal flag is used only by the verified XLSX byte writer. */
export function buildSheetJS(
  input: WorkbookFile,
  options: AdapterOptions,
  preserveFreeze: boolean,
  reportOnly = false,
  discardCellsOnTextBlock = false,
): {
  workbook: SheetJSWorkbook;
  report: CompatibilityReport;
  xlsxTextTooLong?: { sheetName: string; address: string };
} {
  const prior = input.extensions?.['opensheet.importReport'] as CompatibilityReport | undefined;
  validatePriorReport(prior);
  const reader = new FileReader(input),
    snapshot = {
      ...input,
      sheets: input.sheets.map((sh) => reader.sheetData(sh.id)),
      sheetOrder: input.sheets.map((sh) => sh.id),
      styles: Object.fromEntries(input.styles.map((style, i) => [`s${i}`, style])),
    },
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
  const normalize = createPrecedenceNormalizer();
  let xlsxTextTooLong: { sheetName: string; address: string } | undefined;
  for (const id of snapshot.sheetOrder) {
    const sh = snapshot.sheets.find((s) => s.id === id)!;
    let ws: Record<string, any> = {};
    let maxRow = 0,
      maxCol = 0;
    const rows = new Map(sh.rowOrder.map((id, i) => [id, i])),
      cols = new Map(sh.columnOrder.map((id, i) => [id, i]));
    for (const c of reader.cells(id)) {
      const row = rows.get(c.rowId)!,
        col = cols.get(c.columnId)!;
      if (
        !xlsxTextTooLong &&
        c.input.type === 'string' &&
        c.input.value.length > XLSX_CELL_TEXT_LIMIT
      ) {
        xlsxTextTooLong = { sheetName: sh.name, address: address(row, col) };
        if (discardCellsOnTextBlock) {
          workbook.Sheets = Object.create(null);
          ws = {};
        }
      }
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
      const x = exportCell(c, snapshot.styles, r, id, row, col, normalize);
      if (!reportOnly && !(discardCellsOnTextBlock && xlsxTextTooLong)) ws[address(row, col)] = x;
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
    let zeroHeightCount = 0;
    ws['!rows'] = sh.rowOrder.map((id) => {
      const m = sh.rows[id];
      if (m?.size === 0) zeroHeightCount++;
      return m
        ? { hpt: m.size === undefined ? undefined : (m.size * 72) / 96, hidden: m.hidden }
        : undefined;
    });
    if (zeroHeightCount)
      add(r, {
        code: 'ZERO_ROW_HEIGHT_NOT_EXPORTED',
        feature: 'row-height',
        message: 'Zero row height cannot be written by the XLSX writer; use hidden rows instead',
        action: 'dropped',
        severity: 'warning',
        sheetId: id,
        count: zeroHeightCount,
      });
    let widthCount = 0;
    ws['!cols'] = sh.columnOrder.map((id) => {
      const m = sh.columns[id];
      if (m?.size !== undefined) widthCount++;
      return m ? { wpx: m.size, hidden: m.hidden } : undefined;
    });
    if (widthCount)
      add(r, {
        code: 'EXPORT_WIDTH_APPROXIMATED',
        feature: 'column-width',
        message: 'Pixel column widths are converted to font-dependent spreadsheet widths',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        count: widthCount,
      });
    workbook.SheetNames.push(sh.name);
    workbook.Sheets[sh.name] = ws;
    workbook.Workbook.Sheets.push({ name: sh.name, Hidden: sh.hidden ? 1 : 0 });
    if (unsupportedXlsxSheetName(sh.name))
      add(r, {
        code: 'XLSX_SHEET_NAME_BLOCKED',
        feature: 'sheet-name',
        message: XLSX_SHEET_NAME_MESSAGE,
        action: 'blocked',
        severity: 'error',
        sheetId: id,
      });
    if (hasXlsxEscapeText(sh.name))
      add(r, {
        code: 'XLSX_ESCAPE_SHEET_NAME_INTEROP',
        feature: 'sheet-name',
        message: 'Literal _xHHHH_ sequences in sheet names may change when this XLSX is read',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
      });
    if (hasXlsxUnpairedSurrogate(sh.name))
      add(r, {
        code: 'XLSX_UNPAIRED_SURROGATE_SHEET_NAME',
        feature: 'sheet-name',
        message: 'Unpaired UTF-16 surrogates in sheet names may be replaced when written to XLSX',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
      });
    if (hasXlsxControlText(sh.name))
      add(r, {
        code: 'XLSX_CONTROL_SHEET_NAME_INTEROP',
        feature: 'sheet-name',
        message: 'Control characters in sheet names may change when this XLSX is read',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
      });
    const hasFreeze = sh.freeze.rows > 0 || sh.freeze.columns > 0;
    const freezeDropped =
      hasFreeze &&
      (!preserveFreeze || !canWriteFreeze(sh.freeze, sh.rowOrder.length, sh.columnOrder.length));
    if (sh.filter || freezeDropped)
      add(r, {
        code: 'VIEW_NOT_EXPORTED',
        feature: 'view',
        message:
          !preserveFreeze || (sh.filter && freezeDropped)
            ? 'Filters and frozen panes are preserved only in OpenSheet JSON'
            : sh.filter
              ? 'Filters are preserved only in OpenSheet JSON'
              : 'Frozen panes are preserved only in OpenSheet JSON',
        action: 'dropped',
        severity: 'warning',
        sheetId: id,
      });
  }
  if (prior !== undefined) {
    const carried = { approximated: 0, dropped: 0, blocked: 0 };
    for (const issue of prior.issues) {
      if (issue.action === 'preserved') continue;
      const before = r.summary[issue.action];
      add(r, issue);
      carried[issue.action] += r.summary[issue.action] - before;
    }
    for (const action of ['approximated', 'dropped', 'blocked'] as const) {
      const missing = (prior.summary?.[action] ?? 0) - carried[action];
      if (Number.isSafeInteger(missing) && missing > 0)
        add(r, {
          code: 'IMPORT_DETAILS_OMITTED',
          feature: 'import-report',
          action,
          severity: action === 'blocked' ? 'error' : 'warning',
          message: 'Additional original import issues; individual locations were not retained',
          count: missing,
        });
    }
  }
  finish(r, options);
  reader.dispose();
  return { workbook, report: r, xlsxTextTooLong };
}
