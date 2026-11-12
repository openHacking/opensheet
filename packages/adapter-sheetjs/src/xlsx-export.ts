import { address, OpenSheetError, type WorkbookFile } from '@opensheetjs/core';
import { buildSheetJS, canWriteFreeze } from './export.js';
import { unsupportedXlsxSheetName, XLSX_SHEET_NAME_MESSAGE } from './sheet-name.js';
import { generatedWorksheetPaths } from './worksheet-losses.js';
import type { AdapterOptions, SheetJSWorkbook } from './types.js';

interface Archive {
  FullPaths: string[];
  FileIndex: Array<{ content?: Uint8Array }>;
}
/** A SheetJS-compatible writer is injected so the adapter does not bundle it. */
export interface XlsxWriter {
  write(
    book: SheetJSWorkbook,
    options: { type: 'array'; bookType: 'xlsx'; compression?: boolean },
  ): unknown;
  CFB: {
    read(data: Uint8Array, options: { type: 'array' }): Archive;
    write(archive: Archive, options: { type: 'array'; fileType: 'zip' }): unknown;
    utils: { cfb_add(archive: Archive, name: string, content: Uint8Array): unknown };
  };
}

function bytes(value: unknown): Uint8Array {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  throw new OpenSheetError('INVALID_ARGUMENT', 'XLSX writer returned invalid bytes');
}

// SheetJS normally writes stored ZIP entries. Do not inflate an unexpectedly compressed
// archive: the byte-size cap would otherwise say nothing about the allocation required.
function isStoredZip(data: Uint8Array): boolean {
  if (data.length < 22) return false;
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  let end = data.length - 22;
  const lower = Math.max(0, data.length - 22 - 65535);
  while (
    end >= lower &&
    (view.getUint32(end, true) !== 0x06054b50 ||
      end + 22 + view.getUint16(end + 20, true) !== data.length)
  )
    end--;
  if (end < lower || view.getUint16(end + 4, true) !== 0 || view.getUint16(end + 6, true) !== 0)
    return false;
  const count = view.getUint16(end + 10, true);
  const size = view.getUint32(end + 12, true);
  let offset = view.getUint32(end + 16, true);
  if (count !== view.getUint16(end + 8, true) || offset + size !== end) return false;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > end || view.getUint32(offset, true) !== 0x02014b50) return false;
    const flags = view.getUint16(offset + 8, true);
    const local = view.getUint32(offset + 42, true);
    const payloadSize = view.getUint32(offset + 20, true);
    if (
      (flags & ~0x800) !== 0 ||
      view.getUint16(offset + 10, true) !== 0 ||
      view.getUint32(offset + 20, true) !== view.getUint32(offset + 24, true)
    )
      return false;
    if (
      local + 30 > end ||
      view.getUint32(local, true) !== 0x04034b50 ||
      view.getUint16(local + 6, true) !== flags ||
      view.getUint16(local + 8, true) !== 0 ||
      view.getUint32(local + 18, true) !== payloadSize ||
      view.getUint32(local + 22, true) !== payloadSize ||
      local +
        30 +
        view.getUint16(local + 26, true) +
        view.getUint16(local + 28, true) +
        payloadSize >
        end - size
    )
      return false;
    offset +=
      46 +
      view.getUint16(offset + 28, true) +
      view.getUint16(offset + 30, true) +
      view.getUint16(offset + 32, true);
  }
  return offset === end;
}

function paneXml(rows: number, columns: number): string {
  const active = rows && columns ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  const pane = `<pane${columns ? ` xSplit="${columns}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${address(rows, columns)}" activePane="${active}" state="frozen"/>`;
  const selections =
    rows && columns
      ? '<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight"/>'
      : `<selection pane="${active}"/>`;
  return `<sheetViews><sheetView workbookViewId="0">${pane}${selections}</sheetView></sheetViews>`;
}

/** Write XLSX bytes, preserving representable frozen panes when the generated ZIP is verifiable. */
export function writeXlsx(input: WorkbookFile, writer: XlsxWriter, options: AdapterOptions = {}) {
  if (input.sheets.some((sheet) => unsupportedXlsxSheetName(sheet.name))) {
    // Preserve the complete compatibility report without retaining a second,
    // fully materialized copy of all cells in a SheetJS worksheet.
    const rejected = buildSheetJS(input, options, true, true);
    throw new OpenSheetError('UNSUPPORTED_FEATURE', XLSX_SHEET_NAME_MESSAGE, rejected.report);
  }
  const prepared = buildSheetJS(input, options, true, false, true);
  if (prepared.xlsxTextTooLong)
    throw new OpenSheetError(
      'UNSUPPORTED_FEATURE',
      `XLSX text cells must not exceed 32,767 UTF-16 code units; shorten the affected cells before exporting. First location: ${prepared.xlsxTextTooLong.sheetName}!${prepared.xlsxTextTooLong.address}`,
      prepared.report,
    );
  const original = bytes(writer.write(prepared.workbook, { type: 'array', bookType: 'xlsx' }));
  const freezes = new Map<string, { rows: number; columns: number }>();
  for (const sheet of input.sheets) {
    const rows = sheet.rows.order.reduce((sum, [, count]) => sum + count, 0);
    const columns = sheet.columns.order.reduce((sum, [, count]) => sum + count, 0);
    if ((sheet.freeze.rows || sheet.freeze.columns) && canWriteFreeze(sheet.freeze, rows, columns))
      freezes.set(sheet.name, sheet.freeze);
  }
  // Keep the repacking budget unchanged. A compressed fallback can still fit the
  // playground's input limit, but freezes that require patching remain losses.
  if (original.byteLength > 20 * 1024 * 1024) {
    const fallback = freezes.size
      ? buildSheetJS(input, options, false, true).report
      : prepared.report;
    const compressed = bytes(
      writer.write(prepared.workbook, { type: 'array', bookType: 'xlsx', compression: true }),
    );
    return {
      bytes: compressed.byteLength < original.byteLength ? compressed : original,
      report: fallback,
    };
  }
  if (!freezes.size) return { bytes: original, report: prepared.report };
  if (!isStoredZip(original))
    return { bytes: original, report: buildSheetJS(input, options, false, true).report };
  const archive = writer.CFB.read(original, { type: 'array' });
  if (
    !archive ||
    !Array.isArray(archive.FullPaths) ||
    !Array.isArray(archive.FileIndex) ||
    archive.FullPaths.length !== archive.FileIndex.length
  )
    throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid generated XLSX archive');
  const files: NonNullable<SheetJSWorkbook['files']> = Object.create(null);
  for (let i = 0; i < archive.FullPaths.length; i++) {
    const path = archive.FullPaths[i];
    if (!path.startsWith('Root Entry/'))
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid generated XLSX path');
    const name = path.slice('Root Entry/'.length);
    if (!name || name === '\u0001Sh33tJ5') continue;
    if (Object.hasOwn(files, name))
      throw new OpenSheetError('INVALID_ARGUMENT', 'Duplicate generated XLSX path');
    files[name] = archive.FileIndex[i];
  }
  const mappings = generatedWorksheetPaths(files, prepared.workbook.SheetNames);
  const needle = '<sheetViews><sheetView workbookViewId="0"/></sheetViews>';
  let patched = 0;
  for (const [path, name] of mappings) {
    const freeze = freezes.get(name);
    if (!freeze) continue;
    const content = files[path]?.content;
    if (!(content instanceof Uint8Array))
      throw new OpenSheetError('INVALID_ARGUMENT', 'Missing generated worksheet');
    let xml: string;
    try {
      xml = new TextDecoder('utf-8', { fatal: true }).decode(content);
    } catch {
      throw new OpenSheetError('INVALID_ARGUMENT', 'Invalid generated worksheet encoding');
    }
    const index = xml.indexOf(needle);
    if (index < 0 || xml.indexOf(needle, index + needle.length) >= 0)
      throw new OpenSheetError('INVALID_ARGUMENT', 'Unexpected generated worksheet view');
    const result =
      xml.slice(0, index) + paneXml(freeze.rows, freeze.columns) + xml.slice(index + needle.length);
    writer.CFB.utils.cfb_add(archive, `/${path}`, new TextEncoder().encode(result));
    patched++;
  }
  if (patched !== freezes.size)
    throw new OpenSheetError('INVALID_ARGUMENT', 'Missing generated worksheet view');
  return {
    bytes: bytes(writer.CFB.write(archive, { type: 'array', fileType: 'zip' })),
    report: prepared.report,
  };
}
