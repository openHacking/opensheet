import { contains, intersects, OpenSheetError, RangeReader } from '@opensheetjs/core';
import type { ExportOptions } from './types.js';
export function prepareRange(book: RangeReader, config: ExportOptions) {
  const sheet = book.getSheetById(config.sheetId);
  if (!sheet) throw new Error('Unknown sheet');
  const range = config.range ? sheet.range(config.range) : sheet.getUsedRange();
  const rows = range.getDisplayValues(),
    raw = range.getValues(),
    b = range.bounds;
  const data = book.sheetData(sheet.id),
    report: string[] = [],
    requiredPackages: string[] = [];
  const opt = config.options ?? {};
  if (data.merges.some((m) => intersects(m, b) && !contains(b, m)))
    throw new OpenSheetError(
      'INVALID_RANGE',
      'Export the entire merged range, or unmerge it first',
    );
  return { book, sheet, rows, raw, b, data, report, requiredPackages, opt };
}
export type PreparedRange = ReturnType<typeof prepareRange>;
