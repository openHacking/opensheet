import {
  address,
  type CellRecord,
  OpenSheetError,
  type Scalar,
  type CellStyle,
} from '@opensheetjs/core';
import { add } from './report.js';
import type { CompatibilityReport, Issue } from './types.js';
export const errors: Record<number, string> = {
  0: '#NULL!',
  7: '#DIV/0!',
  15: '#VALUE!',
  23: '#REF!',
  29: '#NAME?',
  36: '#NUM!',
  42: '#N/A',
  43: '#GETTING_DATA',
};
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

export type CellIssue = (
  code: string,
  feature: string,
  message: string,
  action?: Issue['action'],
  address?: string,
) => void;
export function importCell(
  c: any,
  row: number,
  col: number,
  dateSystem: '1900' | '1904',
  issue: CellIssue,
) {
  let readOnly = false;
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
      value = { type: 'number', value: dateSerial(c.v, dateSystem) };
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
    rowId: row,
    columnId: col,
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
    issue('RICH_TEXT_FLATTENED', 'rich-text', 'Rich text flattened to plain text', 'approximated');
  if (c.F || c.D) {
    issue(
      'ARRAY_FORMULA',
      'array-formula',
      'Array formula workbook is read-only in this release',
      'blocked',
    );
    readOnly = true;
  }
  return { record, readOnly };
}

export function exportCell(
  c: CellRecord,
  styles: Record<string, CellStyle>,
  r: CompatibilityReport,
  id: string,
  row: number,
  col: number,
): Record<string, any> {
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
    const style = styles[c.styleId];
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
  return x;
}
