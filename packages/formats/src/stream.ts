import {
  RangeReader,
  address,
  contains,
  intersects,
  key,
  parseRange,
  scalarValue,
  OpenSheetError,
  type Workbook,
} from '@opensheetjs/core';
import { renderDelimited, validateDelimitedOptions } from './generators/delimited.js';
import { renderHtml } from './generators/html.js';
import { renderLatex } from './generators/latex.js';
import { renderMarkdown } from './generators/markdown.js';
import type { PreparedRange } from './prepare.js';
import type { DelimitedExportReport, ExportOptions } from './types.js';
/** Bounded input batches; consumers choose their own file/Blob sink. */
export async function* streamExport(
  book: Workbook,
  config: ExportOptions,
  options: { signal?: AbortSignal; onReport?: (report: DelimitedExportReport) => void } = {},
): AsyncGenerator<string> {
  // A stream is consumed across multiple turns; keep its policy and callbacks stable.
  config = { ...config, options: config.options ? { ...config.options } : undefined };
  options = { ...options };
  options.signal?.throwIfAborted();
  if (config.format === 'csv' || config.format === 'tsv') validateDelimitedOptions(config);
  const revision = book.revision;
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
  let sanitized = 0;
  let preservedFormulas = 0;
  const preserve =
    (config.format === 'csv' || config.format === 'tsv') && config.options?.formulas === 'preserve';
  const sample: string[] = [];
  if (config.format === 'html') yield '<table>\n';
  options.signal?.throwIfAborted();
  if (book.revision !== revision)
    throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed during streaming export');
  for await (const batch of book.streamRange(sheet.id, bounds, { ...options, values: !preserve })) {
    if (book.revision !== revision)
      throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed during streaming export');
    const reader = new RangeReader(book.metadata, sheet.id, batch.cells);
    try {
      const source = reader.getSheetById(sheet.id)!;
      const rows: string[][] = [],
        raw = [];
      for (let row = batch.range.startRow; row < batch.range.endRow; row++) {
        const line: string[] = [],
          values = [];
        for (let col = bounds.startColumn; col < bounds.endColumn; col++) {
          const k = key(data.rowOrder[row], data.columnOrder[col]);
          line.push(batch.display[k] ?? '');
          const input = batch.cells[k]?.input;
          const calculated = batch.calculated[k] ?? null;
          if (
            !preserve &&
            calculated &&
            typeof calculated === 'object' &&
            calculated.error === '#LIMIT!'
          )
            throw new OpenSheetError('LIMIT_EXCEEDED', 'Cannot export an incomplete calculation');
          values.push(
            preserve ? (input && input.type !== 'formula' ? scalarValue(input) : null) : calculated,
          );
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
        opt: {
          ...config.options,
          header: first && config.options?.header !== false,
          bom: first && config.options?.bom,
        },
      };
      if (config.format === 'csv' || config.format === 'tsv') {
        const result = renderDelimited(
          prepared,
          { ...config, options: { ...config.options, bom: first && config.options?.bom } },
          options.onReport
            ? (row, column) => {
                sanitized++;
                if (sample.length < 32)
                  sample.push(address(batch.range.startRow + row, bounds.startColumn + column));
              }
            : undefined,
          options.onReport && config.options?.formulas === 'preserve'
            ? () => {
                preservedFormulas++;
              }
            : undefined,
        );
        yield (first ? '' : '\r\n') + result.text;
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
  if ((config.format === 'csv' || config.format === 'tsv') && options.onReport) {
    options.signal?.throwIfAborted();
    if (book.revision !== revision)
      throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed during streaming export');
    options.onReport({
      format: config.format,
      revision,
      sanitized,
      sample,
      truncated: sanitized > sample.length,
      ...(config.options?.formulas === 'preserve' ? { preservedFormulas } : {}),
    });
  }
}
