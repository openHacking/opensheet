import {
  SnapshotReader,
  contains,
  intersects,
  key,
  parseRange,
  OpenSheetError,
  type Workbook,
} from '@opensheetjs/core';
import { renderDelimited } from './generators/delimited.js';
import { renderHtml } from './generators/html.js';
import { renderLatex } from './generators/latex.js';
import { renderMarkdown } from './generators/markdown.js';
import type { PreparedRange } from './prepare.js';
import type { ExportOptions } from './types.js';
/** Bounded input batches; consumers choose their own file/Blob sink. */
export async function* streamExport(
  book: Workbook,
  config: ExportOptions,
  options: { signal?: AbortSignal } = {},
): AsyncGenerator<string> {
  const sheet = book.getSheetById(config.sheetId);
  if (!sheet) throw new OpenSheetError('INVALID_RANGE', 'Unknown sheet');
  const bounds = config.range ? parseRange(config.range) : sheet.getUsedRange().bounds,
    data = book.sheetData(sheet.id);
  if (data.merges.some((m) => intersects(m, bounds) && !contains(bounds, m)))
    throw new OpenSheetError(
      'INVALID_RANGE',
      'Export the entire merged range, or unmerge it first',
    );
  let first = true;
  if (config.format === 'html') yield '<table>\n';
  for await (const batch of book.streamRange(sheet.id, bounds, options)) {
    const snapshot = {
      ...book.metadata,
      sheets: book.metadata.sheets.map((s) =>
        s.id === sheet.id ? { ...s, cells: batch.cells } : s,
      ),
    };
    const reader = new SnapshotReader(snapshot),
      source = reader.getSheetById(sheet.id)!;
    const rows: string[][] = [],
      raw = [];
    for (let row = batch.range.startRow; row < batch.range.endRow; row++) {
      const line: string[] = [],
        values = [];
      for (let col = bounds.startColumn; col < bounds.endColumn; col++) {
        const k = key(data.rowOrder[row], data.columnOrder[col]);
        line.push(batch.display[k] ?? '');
        values.push(batch.calculated[k] ?? null);
      }
      rows.push(line);
      raw.push(values);
    }
    const prepared: PreparedRange = {
      book: reader,
      sheet: source,
      rows,
      raw,
      b: { ...bounds, startRow: batch.range.startRow },
      data,
      report: [],
      requiredPackages: [],
      opt: { ...config.options, header: first && config.options?.header !== false },
    };
    try {
      if (config.format === 'csv' || config.format === 'tsv') {
        const result = renderDelimited(prepared, {
          ...config,
          options: { ...config.options, bom: first && config.options?.bom },
        });
        yield (first ? '' : '\n') + result.text;
      } else if (config.format === 'markdown') {
        const lines = renderMarkdown(prepared).text.split('\n');
        if (!first) lines.splice(1, 1);
        yield (first ? '' : '\n') + lines.join('\n');
      } else if (config.format === 'html')
        yield renderHtml(prepared).text.slice('<table>\n'.length, -'\n</table>'.length) + '\n';
      else if (config.format === 'latex') {
        const lines = renderLatex(prepared).text.split('\n'),
          begin = lines.findIndex((s) => s.startsWith('\\begin{tabular}'));
        if (first) {
          if (
            data.merges.some((m) => intersects(m, bounds) && m.endRow - m.startRow > 1) &&
            !lines.some((s) => s.includes('usepackage{multirow}'))
          )
            lines.unshift('% Requires \\usepackage{multirow}');
          const footer = lines.length - 2;
          yield lines.slice(0, footer).join('\n') + '\n';
        } else yield lines.slice(begin + 2, -2).join('\n') + '\n';
      } else throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown export format');
    } finally {
      reader.dispose();
    }
    first = false;
  }
  if (config.format === 'html') yield '</table>';
  if (config.format === 'latex')
    yield `${config.options?.booktabs !== false ? '\\bottomrule' : '\\hline'}\n\\end{tabular}`;
}
