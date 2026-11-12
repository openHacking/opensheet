import { contains, intersects, OpenSheetError, RangeReader, scalarValue } from '@opensheetjs/core';
import type { ExportOptions } from './types.js';
export function prepareRange(book: RangeReader, config: ExportOptions) {
  const sheet = book.getSheetById(config.sheetId);
  if (!sheet) throw new Error('Unknown sheet');
  const range = config.range ? sheet.range(config.range) : sheet.getUsedRange();
  const b = range.bounds;
  const data = book.sheetData(sheet.id),
    report: string[] = [],
    requiredPackages: string[] = [];
  const opt = config.options ?? {};
  if (data.merges.some((m) => intersects(m, b) && !contains(b, m)))
    throw new OpenSheetError(
      'INVALID_RANGE',
      'Export the entire merged range, or unmerge it first',
    );
  const preserve =
    (config.format === 'csv' || config.format === 'tsv') && opt.formulas === 'preserve';
  const raw = preserve
    ? Array.from({ length: b.endRow - b.startRow }, (_, row) =>
        Array.from({ length: b.endColumn - b.startColumn }, (_, column) => {
          const input = book.getCell(sheet.id, b.startRow + row, b.startColumn + column)?.input;
          return input && input.type !== 'formula' ? scalarValue(input) : null;
        }),
      )
    : range.getValues();
  if (
    !preserve &&
    raw.some((row) =>
      row.some((value) => value && typeof value === 'object' && value.error === '#LIMIT!'),
    )
  )
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Cannot export an incomplete calculation');
  const rows = preserve
    ? []
    : raw.map((row, r) =>
        row.map((value, c) => book.format(sheet.id, b.startRow + r, b.startColumn + c, value)),
      );
  return { book, sheet, rows, raw, b, data, report, requiredPackages, opt };
}
export type PreparedRange = ReturnType<typeof prepareRange>;
