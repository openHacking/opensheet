import { stringifyDelimited } from '../delimited.js';
import { OpenSheetError } from '@opensheetjs/core';
import type { PreparedRange } from '../prepare.js';
import type { ExportOptions, ExportResult } from '../types.js';
export function validateDelimitedOptions(config: ExportOptions) {
  const mode = config.options?.formulas;
  if (mode !== undefined && mode !== 'values' && mode !== 'preserve')
    throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown delimited formula mode');
  if (mode === 'preserve' && config.options?.safe !== false)
    throw new OpenSheetError('INVALID_ARGUMENT', 'Formula export requires explicit safe: false');
}
export function renderDelimited(
  prepared: PreparedRange,
  config: ExportOptions,
  onSanitized?: (row: number, column: number) => void,
  onFormula?: () => void,
): ExportResult {
  const { raw, data, report, requiredPackages, opt, book, sheet, b } = prepared;
  validateDelimitedOptions(config);
  let formulas = 0;

  const result = stringifyDelimited(
    raw.map((row, r) =>
      row.map((v, c) => {
        if (opt.formulas === 'preserve') {
          const input = book.getCell(sheet.id, b.startRow + r, b.startColumn + c)?.input;
          if (input?.type === 'formula') {
            formulas++;
            onFormula?.();
            return input.expression.startsWith('=') ? input.expression : `=${input.expression}`;
          }
        }
        return v && typeof v === 'object' ? v.error : v;
      }),
    ),
    config.format === 'csv' ? ',' : '\t',
    opt.safe !== false,
    onSanitized,
  );
  if (result.sanitized)
    report.push(
      `${result.sanitized} text cells prefixed with an apostrophe for safer spreadsheet viewing; import behavior varies.`,
    );
  if (opt.formulas === 'preserve')
    report.push(
      `${formulas} formulas exported as formulas; raw text is not escaped and spreadsheet applications may execute formulas. References are not rebased and may depend on omitted cells or sheets.`,
    );
  if (data.merges.length) report.push('Merged cells are flattened in delimited text.');
  return { text: (opt.bom ? '\uFEFF' : '') + result.text, report, requiredPackages };
}
