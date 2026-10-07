import type { PreparedRange } from '../prepare.js';
import type { ExportResult } from '../types.js';
export function renderMarkdown(prepared: PreparedRange): ExportResult {
  const { rows, data, report, requiredPackages } = prepared;

  if (data.merges.length) report.push('Markdown has no merged cells; covered cells remain blank.');
  const escaped = rows.map((row) =>
    row.map((v) => v.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')),
  );
  const line = (r: string[]) => '| ' + r.join(' | ') + ' |';
  return {
    text: [line(escaped[0]), line(escaped[0].map(() => '---')), ...escaped.slice(1).map(line)].join(
      '\n',
    ),
    report,
    requiredPackages,
  };
}
