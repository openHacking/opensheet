import {
  address,
  isSafeLinkTarget,
  isSupportedNumberFormat,
  formulaHelp,
  LIMITS,
  type CellRecord,
  OpenSheetError,
  type Scalar,
  type CellStyle,
} from '@opensheetjs/core';
import { add } from './report.js';
import {
  createPrecedenceNormalizer,
  mayDifferInPowerPrecedence,
  toXlsxFormula,
} from './formula.js';
import type { CompatibilityReport, Issue } from './types.js';
const functionSupport = new Map<string, boolean>();
export const XLSX_CELL_TEXT_LIMIT = 32767;
/** SheetJS interprets literal OOXML escape-shaped text while reading XLSX. */
export function hasXlsxEscapeText(value: unknown): boolean {
  return typeof value === 'string' && /_x[0-9a-f]{4}_/i.test(value);
}
export function hasXlsxControlText(value: unknown): boolean {
  return typeof value === 'string' && /[\u0000-\u0008\u000b-\u001f\ufffe\uffff]/.test(value);
}
export function hasXlsxUnpairedSurrogate(value: unknown): boolean {
  return (
    typeof value === 'string' &&
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(value)
  );
}
function hasBracketedReference(expression: string): boolean {
  if (!expression.includes('[')) return false;
  let quoted = false;
  let open = false;
  let content = false;
  for (let i = 0; i < expression.length; i++) {
    const char = expression[i];
    if (char === '"') {
      if (quoted && expression[i + 1] === '"') i++;
      else quoted = !quoted;
    } else if (!quoted) {
      if (char === '[') {
        if (open) content = true;
        else open = true;
      } else if (char === ']') {
        if (open && content) return true;
        open = false;
        content = false;
      } else if (open) content = true;
    }
  }
  return false;
}
function* functionNames(expression: string): Generator<{ name: string; start: number }> {
  const identifier = /[A-Za-z_][A-Za-z0-9_.]*/y;
  for (let i = 0; i < expression.length; ) {
    const char = expression[i];
    if (char === '"' || char === "'") {
      i++;
      while (i < expression.length) {
        const end = expression.indexOf(char, i);
        if (end < 0) {
          i = expression.length;
          break;
        }
        i = end + 1;
        if (expression[i] !== char) break;
        i++;
      }
      continue;
    }
    identifier.lastIndex = i;
    const match = identifier.exec(expression);
    if (!match) {
      i++;
      continue;
    }
    i = identifier.lastIndex;
    while (i < expression.length && /\s/.test(expression[i])) i++;
    if (expression[i] !== '(') continue;
    yield { name: match[0].toUpperCase().replace(/^_XLFN\.(?:_XLWS\.)?/, ''), start: match.index };
  }
}
function unsupportedFunction(expression: string): string | undefined {
  for (const { name } of functionNames(expression)) {
    let supported = functionSupport.get(name);
    if (supported === undefined) {
      supported = formulaHelp(`=${name}(`).call?.definition.name === name;
      if (name.length <= 128 && functionSupport.size < 256) functionSupport.set(name, supported);
    }
    if (!supported) return name;
  }
}
function lookupInteropRisk(expression: string): boolean {
  if (!/(?:MATCH|VLOOKUP)\s*\(/i.test(expression)) return false;
  for (const { name, start } of functionNames(expression)) {
    if (name !== 'MATCH' && name !== 'VLOOKUP') continue;
    const previous = start > 0 ? expression.codePointAt(start - 1)! : 0;
    const code =
      previous >= 0xdc00 && previous <= 0xdfff && start > 1
        ? expression.codePointAt(start - 2)!
        : previous;
    if (code && /[\p{L}\p{N}\p{M}_.\\]/u.test(String.fromCodePoint(code))) continue;
    return true;
  }
  return false;
}
const lookupInteropMessage =
  'MATCH and VLOOKUP may handle lookup errors, blank cells or wildcards differently across spreadsheet applications';

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
  let local: number;
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime()))
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid date');
    const wall = new Date(0);
    wall.setUTCFullYear(value.getUTCFullYear(), value.getUTCMonth(), value.getUTCDate());
    wall.setUTCHours(
      value.getUTCHours(),
      value.getUTCMinutes(),
      value.getUTCSeconds(),
      value.getUTCMilliseconds(),
    );
    local = wall.getTime();
  } else {
    const match =
      typeof value === 'string' &&
      /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(Z|[+-]\d{2}:\d{2})?)?$/.exec(
        value,
      );
    if (!match) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid ISO date');
    const [, y, m, day, h = '0', minute = '0', second = '0', fraction = '', zone] = match;
    const year = Number(y),
      month = Number(m),
      date = Number(day),
      hours = Number(h),
      minutes = Number(minute),
      seconds = Number(second);
    const wall = new Date(0);
    wall.setUTCFullYear(year, month - 1, date);
    wall.setUTCHours(hours, minutes, seconds, Number(fraction.padEnd(3, '0')));
    if (
      month < 1 ||
      month > 12 ||
      date < 1 ||
      wall.getUTCFullYear() !== year ||
      wall.getUTCMonth() !== month - 1 ||
      wall.getUTCDate() !== date ||
      hours > 23 ||
      minutes > 59 ||
      seconds > 59
    )
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid ISO date');
    local = wall.getTime();
    if (zone && zone !== 'Z') {
      const offsetHours = Number(zone.slice(1, 3)),
        offsetMinutes = Number(zone.slice(4, 6));
      if (offsetHours > 23 || offsetMinutes > 59)
        throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid ISO date offset');
      local -= (zone[0] === '+' ? 1 : -1) * (offsetHours * 60 + offsetMinutes) * 60000;
    }
  }
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
  normalize: ReturnType<typeof createPrecedenceNormalizer>,
  trustedLinkTarget?: string,
) {
  let formula = c.f ? String(c.f) : undefined;
  if (formula && formula.length > LIMITS.text)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Formula exceeds text limit');
  if (formula && mayDifferInPowerPrecedence(formula)) {
    const normalized = normalize(formula, 'spreadsheet');
    if (normalized) {
      issue(
        'IMPORT_FORMULA_PRECEDENCE_NORMALIZED',
        'formula-precedence',
        'Formula parentheses preserve spreadsheet exponent precedence',
        'preserved',
        address(row, col),
      );
      formula = normalized;
    } else
      issue(
        'IMPORT_FORMULA_POWER_PRECEDENCE',
        'formula-precedence',
        'Exponentiation and negation may calculate differently in OpenSheet',
        'approximated',
        address(row, col),
      );
  }
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
      if (c.v === undefined && c.f) {
        value = { type: 'blank' };
        break;
      }
      if (typeof c.v !== 'boolean')
        throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid boolean cell');
      value = { type: 'boolean', value: c.v };
      break;
    case 'e':
      if (c.v === undefined && c.f) {
        value = { type: 'blank' };
        break;
      }
      if (!Number.isSafeInteger(c.v))
        throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid error cell');
      if (!Object.hasOwn(errors, c.v))
        issue(
          'ERROR_CODE_APPROXIMATED',
          'error',
          'Unknown spreadsheet error converted to #VALUE!',
          'approximated',
          address(row, col),
        );
      value = { type: 'error', code: Object.hasOwn(errors, c.v) ? errors[c.v] : '#VALUE!' };
      break;
    case 'd':
      if (c.v === undefined && c.f) {
        value = { type: 'blank' };
        break;
      }
      value = { type: 'number', value: dateSerial(c.v, dateSystem) };
      issue(
        'DATE_NORMALIZED',
        'date',
        'Date converted to spreadsheet serial',
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
    input: formula ? { type: 'formula', expression: formula } : value,
  };
  if (formula && hasBracketedReference(formula))
    issue(
      'UNSUPPORTED_EXTERNAL_OR_STRUCTURED_REFERENCE',
      'formula-reference',
      'External workbook and structured table references are not evaluated',
      'dropped',
      address(row, col),
    );
  if (formula && lookupInteropRisk(formula))
    issue(
      'XLSX_LOOKUP_IMPORT_INTEROP',
      'lookup-semantics',
      lookupInteropMessage,
      'approximated',
      address(row, col),
    );
  if (formula) {
    const name = unsupportedFunction(formula);
    if (name)
      issue(
        'UNSUPPORTED_FORMULA_FUNCTION',
        'formula-function',
        'Formula uses a function that OpenSheet does not support',
        'dropped',
        address(row, col),
      );
  }
  // SheetJS uses a stub with a synthetic zero for formulas lacking a cached value.
  if (c.f && c.t !== 'z' && c.v !== undefined) record.cached = value;
  if (c.z) {
    record.numberFormat = String(c.z);
    if ((c.t === 'n' || c.t === 'd' || c.f) && !isSupportedNumberFormat(record.numberFormat))
      issue(
        'NUMBER_FORMAT_APPROXIMATED',
        'number-format',
        'Number format is retained but may display approximately',
        'approximated',
        address(row, col),
      );
  } else if (c.t === 'd') record.numberFormat = 'yyyy-mm-dd';
  if (c.l?.Target) {
    const target = trustedLinkTarget ?? c.l.Target;
    if (isSafeLinkTarget(target)) {
      record.link = {
        target,
        ...(c.l.Tooltip !== undefined ? { tooltip: String(c.l.Tooltip) } : {}),
      };
      if (trustedLinkTarget === undefined && /[^\x00-\x7f]/.test(target))
        issue(
          'XLSX_LINK_URL_IMPORT_UNCERTAIN',
          'link',
          'A non-ASCII hyperlink target may have been decoded incorrectly; verify its destination',
          'approximated',
          address(row, col),
        );
    } else
      issue(
        'UNSAFE_LINK',
        'link',
        'Unsafe or malformed hyperlink omitted',
        'dropped',
        address(row, col),
      );
  }
  if (Array.isArray(c.c)) record.note = c.c.map((x: any) => String(x.t ?? '')).join('\n');
  // SheetJS exposes the default no-fill style even on unstyled XLSX cells.
  if (c.s && Object.keys(c.s).some((key) => key !== 'patternType' || c.s.patternType !== 'none'))
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
  normalize: ReturnType<typeof createPrecedenceNormalizer>,
): Record<string, any> {
  const x: Record<string, any> = {};
  const v = c.input;
  if (v.type === 'formula') {
    const risky = mayDifferInPowerPrecedence(v.expression);
    const normalized = risky ? normalize(v.expression, 'native') : undefined;
    x.f = toXlsxFormula(normalized ?? v.expression);
    x.t = 'n';
    if (lookupInteropRisk(v.expression))
      add(r, {
        code: 'XLSX_LOOKUP_EXPORT_INTEROP',
        feature: 'lookup-semantics',
        message: lookupInteropMessage,
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
    if (x.f.length >= 8192)
      add(r, {
        code: 'XLSX_FORMULA_LENGTH_INTEROP',
        feature: 'formula',
        message:
          'Formulas of 8,192 or more UTF-16 code units (excluding the leading equals sign) may fail in some spreadsheet applications',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
    if (risky)
      add(r, {
        code: normalized
          ? 'EXPORT_FORMULA_PRECEDENCE_NORMALIZED'
          : 'EXPORT_FORMULA_POWER_PRECEDENCE',
        feature: 'formula-precedence',
        message: normalized
          ? 'Formula parentheses preserve OpenSheet exponent precedence'
          : 'Exponentiation and negation may calculate differently in the target application',
        action: normalized ? 'preserved' : 'approximated',
        severity: normalized ? 'info' : 'warning',
        sheetId: id,
        address: address(row, col),
      });
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
    const mapped = Object.entries(errors).find(([, name]) => name === v.code);
    if (!mapped)
      add(r, {
        code: 'ERROR_CODE_APPROXIMATED',
        feature: 'error',
        message: 'Unsupported spreadsheet error converted to #VALUE!',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
    x.v = Number(mapped?.[0] ?? 15);
  } else {
    x.t = v.type === 'number' ? 'n' : v.type === 'boolean' ? 'b' : 's';
    x.v = v.value;
    if (v.type === 'string' && v.value.length > XLSX_CELL_TEXT_LIMIT)
      add(r, {
        code: 'XLSX_TEXT_LENGTH_BLOCKED',
        feature: 'text',
        message: 'XLSX text cells must not exceed 32,767 UTF-16 code units',
        action: 'blocked',
        severity: 'error',
        sheetId: id,
        address: address(row, col),
      });
    if (v.type === 'string' && v.value === '')
      add(r, {
        code: 'EMPTY_TEXT_INTEROP',
        feature: 'empty-text',
        message: 'Empty text may be lost when another spreadsheet application saves the file',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
  }
  if (c.numberFormat) x.z = c.numberFormat;
  if (c.link && isSafeLinkTarget(c.link.target)) {
    let target = c.link.target;
    if (/[^\x00-\x7f]/.test(target)) {
      // The pinned SheetJS reader decodes non-ASCII relationship targets as
      // bytes. URL canonicalization yields the same browser destination in ASCII.
      const normalized = new URL(target).href;
      if (normalized.length <= 4096) {
        target = normalized;
        add(r, {
          code: 'XLSX_LINK_URL_NORMALIZED',
          feature: 'link',
          message:
            'Unicode link targets are normalized to URL-encoded ASCII for XLSX; their text form changes',
          action: 'approximated',
          severity: 'warning',
          sheetId: id,
          address: address(row, col),
        });
      } else {
        add(r, {
          code: 'XLSX_LINK_URL_UNSAFE_LENGTH',
          feature: 'link',
          message:
            'Unicode link target cannot be normalized within the supported length and may be corrupted on import',
          action: 'approximated',
          severity: 'warning',
          sheetId: id,
          address: address(row, col),
        });
      }
    }
    x.l = { Target: target, Tooltip: c.link.tooltip };
    // SheetJS otherwise omits a blank cell before collecting its hyperlink.
    // Explicit General emits a blank cell without converting it to empty text.
    if (v.type === 'blank') {
      if (x.z === undefined) x.z = 'General';
      add(r, {
        code: 'BLANK_LINK_INTEROP',
        feature: 'link',
        message:
          'Hyperlinks on blank cells may be lost when another spreadsheet application saves the file; some readers may expose the target as the cell value',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
    }
    if (c.link.tooltip === '')
      add(r, {
        code: 'EMPTY_LINK_TOOLTIP_DROPPED',
        feature: 'link',
        message: 'An explicitly empty link tooltip is omitted by the XLSX writer',
        action: 'dropped',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
    else if (c.link.tooltip !== undefined)
      add(r, {
        code: 'LINK_TOOLTIP_INTEROP',
        feature: 'link',
        message: 'Link tooltips may be removed when another spreadsheet application saves the file',
        action: 'approximated',
        severity: 'warning',
        sheetId: id,
        address: address(row, col),
      });
  }
  if (c.note !== undefined) x.c = [{ a: 'OpenSheet', t: c.note }];
  if (c.styleId) {
    const style = styles[c.styleId];
    if (style.numberFormat && c.numberFormat === undefined) x.z = style.numberFormat;
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
  if (
    hasXlsxEscapeText(x.v) ||
    hasXlsxEscapeText(x.f) ||
    hasXlsxEscapeText(x.z) ||
    hasXlsxEscapeText(c.note) ||
    hasXlsxEscapeText(c.link?.target)
  )
    add(r, {
      code: 'XLSX_ESCAPE_TEXT_INTEROP',
      feature: 'text',
      message:
        'Literal _xHHHH_ sequences may change when this XLSX is read, including in formulas, notes or link targets',
      action: 'approximated',
      severity: 'warning',
      sheetId: id,
      address: address(row, col),
    });
  if (
    hasXlsxUnpairedSurrogate(x.v) ||
    hasXlsxUnpairedSurrogate(x.f) ||
    hasXlsxUnpairedSurrogate(x.z) ||
    hasXlsxUnpairedSurrogate(c.note) ||
    hasXlsxUnpairedSurrogate(c.link?.target) ||
    hasXlsxUnpairedSurrogate(c.link?.tooltip)
  )
    add(r, {
      code: 'XLSX_UNPAIRED_SURROGATE_TEXT',
      feature: 'text',
      message: 'Unpaired UTF-16 surrogates may be replaced when written to XLSX',
      action: 'approximated',
      severity: 'warning',
      sheetId: id,
      address: address(row, col),
    });
  if (
    hasXlsxControlText(x.v) ||
    hasXlsxControlText(x.f) ||
    hasXlsxControlText(x.z) ||
    hasXlsxControlText(c.note) ||
    hasXlsxControlText(c.link?.target) ||
    hasXlsxControlText(c.link?.tooltip)
  )
    add(r, {
      code: 'XLSX_CONTROL_TEXT_INTEROP',
      feature: 'text',
      message:
        'Some spreadsheet applications may display control characters as escape sequences or replace Unicode noncharacters',
      action: 'approximated',
      severity: 'warning',
      sheetId: id,
      address: address(row, col),
    });
  return x;
}
