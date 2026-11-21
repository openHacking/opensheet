import {
  createWorkbookFile,
  LIMITS,
  OpenSheetError,
  address,
  parseAddress,
  validateWorkbookFile,
  encodeFileCells,
  isSafeLinkTarget,
  type CellRecord,
  type WorkbookFile,
} from '@opensheetjs/core';
import { importCell } from './cells.js';
import { add, finish, report } from './report.js';
import { rawHyperlinkTargets, reportWorksheetLosses } from './worksheet-losses.js';
import { createPrecedenceNormalizer } from './formula.js';
import { definedNameIndex, referencesDefinedName } from './defined-name-scan.js';
import type { AdapterOptions, CompatibilityReport, Issue, SheetJSWorkbook } from './types.js';
const hyperlinkRelationship =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
function corroboratedLinkTarget(
  link: any,
  ref: string,
  candidates:
    | {
        links: Map<string, { target: string; byteDecoded: string }>;
        cells: Map<string, { id: string; ref: string; location?: string }>;
      }
    | undefined,
): string | undefined {
  if (!link || typeof link.id !== 'string' || !candidates) return;
  const cell = candidates.cells.get(ref);
  if (!cell || cell.id !== link.id || cell.location !== link.location || link.ref !== cell.ref)
    return;
  const candidate = candidates.links.get(link.id),
    rel = link.Rel;
  if (
    !candidate ||
    !rel ||
    rel.Id !== link.id ||
    rel.Type !== hyperlinkRelationship ||
    rel.TargetMode !== 'External' ||
    typeof rel.Target !== 'string' ||
    (rel.Target !== candidate.target && rel.Target !== candidate.byteDecoded) ||
    (link.location !== undefined && typeof link.location !== 'string')
  )
    return;
  const suffix = link.location ? `#${link.location}` : '';
  if (link.Target !== rel.Target + suffix) return;
  const target = candidate.target + suffix;
  return isSafeLinkTarget(target) ? target : undefined;
}

export function fromSheetJS(
  input: SheetJSWorkbook,
  options: AdapterOptions = {},
): { file: WorkbookFile; report: CompatibilityReport } {
  if (
    !input ||
    !Array.isArray(input.SheetNames) ||
    input.SheetNames.length < 1 ||
    input.SheetNames.length > LIMITS.sheets
  )
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Invalid number of sheets');
  const definedNames = input.Workbook?.Names;
  if (definedNames != null && !Array.isArray(definedNames))
    throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid defined-name collection');
  if (definedNames)
    for (let i = 0; i < definedNames.length; i++) {
      const entry = definedNames[i];
      if (
        !Object.hasOwn(definedNames, i) ||
        !entry ||
        typeof entry !== 'object' ||
        typeof entry.Name !== 'string' ||
        !entry.Name
      )
        throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid defined-name collection');
    }
  const snapshot = createWorkbookFile();
  snapshot.sheets = [];
  snapshot.dateSystem = input.Workbook?.WBProps?.date1904 ? '1904' : '1900';
  const r = report();
  const nameIndex = definedNameIndex(definedNames ?? [], input.SheetNames.length);
  const qualifiedLocalNames = new Map(
    input.SheetNames.flatMap((name, index) => {
      const names = nameIndex.local.get(index);
      return typeof name === 'string' && (names || nameIndex.global.size)
        ? [[name.toUpperCase(), names ?? new Set<string>()] as const]
        : [];
    }),
  );
  if (nameIndex.incomplete)
    add(r, {
      code: 'DEFINED_NAME_SCAN_INCOMPLETE',
      feature: 'defined-name-reference',
      message: 'Some defined names could not be included in the formula reference scan',
      action: 'approximated',
      severity: 'warning',
    });
  const normalize = createPrecedenceNormalizer();
  const freezes = reportWorksheetLosses(input.files, r, input.SheetNames, input.Workbook?.Sheets);
  let rawLinks: ReturnType<typeof rawHyperlinkTargets> | undefined;
  let rawLinksWithAscii: ReturnType<typeof rawHyperlinkTargets> | undefined;
  for (const [index, name] of input.SheetNames.entries()) {
    const ws =
      input.Sheets && typeof input.Sheets === 'object' && Object.hasOwn(input.Sheets, name)
        ? input.Sheets[name]
        : undefined;
    if (!ws || typeof ws !== 'object') {
      const label = typeof name === 'string' && name.length <= 31 ? ` ${JSON.stringify(name)}` : '';
      throw new OpenSheetError(
        'INVALID_ARGUMENT',
        `Missing worksheet${label}; the file may be incomplete or the parser could not read it`,
      );
    }
    let maxRow = 1,
      maxCol = 1;
    const collect = (row: number, col: number, c: any) => {
      if (c === undefined || c === null) return;
      if (row >= LIMITS.rows || col >= LIMITS.columns)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Workbook exceeds preview limits');
      maxRow = Math.max(maxRow, row + 1);
      maxCol = Math.max(maxCol, col + 1);
    };
    const hasDense = Array.isArray(ws) || Object.hasOwn(ws, '!data');
    const dense = Array.isArray(ws) ? ws : hasDense ? ws['!data'] : undefined;
    const visitCells = (visit: (row: number, col: number, cell: any) => void) => {
      if (hasDense) {
        if (!Array.isArray(dense))
          throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid dense worksheet data');
        if (dense.length > LIMITS.rows)
          throw new OpenSheetError('LIMIT_EXCEEDED', 'Dense worksheet too large');
        for (let row = 0; row < dense.length; row++) {
          if (dense[row] === undefined || dense[row] === null) continue;
          if (!Array.isArray(dense[row]))
            throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid dense worksheet row');
          if (dense[row].length > LIMITS.columns)
            throw new OpenSheetError('LIMIT_EXCEEDED', 'Too many columns');
          for (let col = 0; col < dense[row].length; col++) visit(row, col, dense[row][col]);
        }
      } else {
        for (const a of Object.keys(ws)) {
          if (a.startsWith('!')) continue;
          const p = parseAddress(a);
          visit(p.row, p.column, ws[a]);
        }
      }
    };
    visitCells(collect);
    const merges =
      Object.hasOwn(ws, '!merges') && Array.isArray(ws['!merges']) ? ws['!merges'] : [];
    if (merges.length > 10000) throw new OpenSheetError('LIMIT_EXCEEDED', 'Too many merged ranges');
    for (const m of merges) {
      if (!m.s || !m.e) throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid merge');
      maxRow = Math.max(maxRow, m.e.r + 1);
      maxCol = Math.max(maxCol, m.e.c + 1);
    }
    const rowMeta = (Object.hasOwn(ws, '!rows') && ws['!rows']) || [],
      colMeta = (Object.hasOwn(ws, '!cols') && ws['!cols']) || [];
    maxRow = Math.max(maxRow, rowMeta.length);
    maxCol = Math.max(maxCol, colMeta.length);
    const freeze = freezes.get(name);
    if (freeze) {
      maxRow = Math.max(maxRow, freeze.rows + 1);
      maxCol = Math.max(maxCol, freeze.columns + 1);
    }
    const sh = createWorkbookFile({
      sheets: [{ name, rows: Math.max(100, maxRow), columns: Math.max(26, maxCol) }],
    }).sheets[0];
    if (freeze) sh.freeze = freeze;
    sh.hidden = !!input.Workbook?.Sheets?.[index]?.Hidden;
    snapshot.sheets.push(sh);
    const blocks = new Map<string, CellRecord[]>();
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
    visitCells((row, col, c) => {
      if (c === undefined || c === null) return;
      let trustedLinkTarget: string | undefined;
      if (typeof c.l?.Target === 'string' && /[^\x00-\x7f]/.test(c.l.Target)) {
        const asciiRelationship =
          typeof c.l.Rel?.Target === 'string' && !/[^\x00-\x7f]/.test(c.l.Rel.Target);
        const candidates = asciiRelationship
          ? (rawLinksWithAscii ??= rawHyperlinkTargets(
              input.files,
              input.SheetNames,
              input.Workbook?.Sheets,
              true,
            ))
          : (rawLinks ??= rawHyperlinkTargets(
              input.files,
              input.SheetNames,
              input.Workbook?.Sheets,
            ));
        trustedLinkTarget = corroboratedLinkTarget(c.l, address(row, col), candidates.get(name));
        if (trustedLinkTarget !== undefined && trustedLinkTarget !== c.l.Target)
          add(r, {
            code: 'XLSX_LINK_URL_UTF8_RESTORED',
            feature: 'link',
            message: 'Unicode link target restored from a verified XLSX relationship',
            action: 'preserved',
            severity: 'info',
            sheetId: sh.id,
            address: address(row, col),
          });
      }
      const { record, readOnly } = importCell(
        c,
        row,
        col,
        snapshot.dateSystem,
        issue,
        normalize,
        trustedLinkTarget,
      );
      if (
        record.input.type === 'formula' &&
        referencesDefinedName(
          record.input.expression,
          nameIndex.global,
          nameIndex.local.get(index),
          qualifiedLocalNames,
        )
      )
        issue(
          'DEFINED_NAME_REFERENCE_DROPPED',
          'defined-name-reference',
          'Formula references a defined name that is not preserved',
          'dropped',
          address(row, col),
        );
      if (readOnly) snapshot.extensions['opensheet.readOnly'] = true;
      const blockKey = `${Math.floor(row / 64)}/${Math.floor(col / 32)}`;
      if (!blocks.has(blockKey)) blocks.set(blockKey, []);
      blocks.get(blockKey)!.push(record);
      r.summary.exact++;
    });
    sh.blocks = [...blocks].map(([address, cells]) => {
      const [row, column] = address.split('/').map(Number);
      return { row, column, data: encodeFileCells(cells, new Map()) };
    });
    sh.merges = merges.map((m: any) => ({
      startRow: m.s.r,
      startColumn: m.s.c,
      endRow: m.e.r + 1,
      endColumn: m.e.c + 1,
    }));
    rowMeta.forEach((m: any, i: number) => {
      if (m)
        (sh.rows.meta ??= {})[i] = {
          ...(typeof m.hpt === 'number'
            ? { size: (m.hpt * 96) / 72 }
            : typeof m.hpx === 'number'
              ? { size: m.hpx }
              : {}),
          ...(m.hidden ? { hidden: true } : {}),
        };
    });
    colMeta.forEach((m: any, i: number) => {
      if (m) {
        (sh.columns.meta ??= {})[i] = {
          ...(typeof m.wpx === 'number'
            ? { size: m.wpx }
            : typeof m.wch === 'number'
              ? { size: m.wch * 7 + 5 }
              : {}),
          ...(m.hidden ? { hidden: true } : {}),
        };
        if (typeof m.width === 'number' || (typeof m.wch === 'number' && typeof m.wpx !== 'number'))
          issue(
            'WIDTH_APPROXIMATED',
            'column-width',
            'Spreadsheet column width is font-dependent and its pixel width is approximate',
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
  if (definedNames?.length)
    add(r, {
      code: 'NAMES_UNSUPPORTED',
      feature: 'defined-names',
      message: 'Defined names are not preserved',
      action: 'dropped',
      severity: 'warning',
      count: definedNames.length,
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
  return { file: validateWorkbookFile(snapshot), report: r };
}
