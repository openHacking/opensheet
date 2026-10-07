import { stringifyDelimited } from '../delimited.js';
import type { PreparedRange } from '../prepare.js';
import type { ExportOptions, ExportResult } from '../types.js';
export function renderDelimited(prepared: PreparedRange, config: ExportOptions): ExportResult {
  const { raw, data, report, requiredPackages, opt } = prepared;

  const result = stringifyDelimited(
    raw.map((row) => row.map((v) => (v && typeof v === 'object' ? v.error : v))),
    config.format === 'csv' ? ',' : '\t',
    opt.safe !== false,
  );
  if (result.sanitized)
    report.push(`${result.sanitized} text cells escaped to prevent spreadsheet formula injection.`);
  if (data.merges.length) report.push('Merged cells are flattened in delimited text.');
  return { text: (opt.bom ? '\uFEFF' : '') + result.text, report, requiredPackages };
}
