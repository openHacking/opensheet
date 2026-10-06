import {
  Workbook,
  contains,
  intersects,
  OpenSheetError,
  type WorkbookSnapshot,
  type Rect,
  type CellValue,
} from '@opensheetjs/core';
export type Format = 'csv' | 'tsv' | 'markdown' | 'html' | 'latex';
export type ExportOptions = {
  sheetId: string;
  range?: string | Rect;
  format: Format;
  options?: { safe?: boolean; bom?: boolean; booktabs?: boolean; header?: boolean };
};
export const escapeHTML = (text: string) =>
  text.replace(
    /[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]!,
  );
export const escapeLatex = (text: string) =>
  text
    .replace(
      /[&%$#_{}~^\\]/g,
      (c) =>
        ({
          '&': '\\&',
          '%': '\\%',
          $: '\\$',
          '#': '\\#',
          _: '\\_',
          '{': '\\{',
          '}': '\\}',
          '~': '\\textasciitilde{}',
          '^': '\\textasciicircum{}',
          '\\': '\\textbackslash{}',
        })[c]!,
    )
    .replace(/\r?\n/g, ' ');
export function stringifyDelimited(
  rows: unknown[][],
  separator = ',',
  safe = true,
): { text: string; sanitized: number } {
  let sanitized = 0;
  const text = rows
    .map((row) =>
      row
        .map((value) => {
          let s = String(value ?? '');
          if (safe && typeof value === 'string' && /^[\s]*[=+@-]|^[\t\r]/.test(s)) {
            s = "'" + s;
            sanitized++;
          }
          return s.includes(separator) || /["\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        })
        .join(separator),
    )
    .join('\r\n');
  return { text, sanitized };
}
export function parseDelimited(text: string, separator = ','): string[][] {
  if (text.length > 20 * 1024 * 1024) throw new Error('Text file exceeds 20 MiB');
  const rows: string[][] = [];
  let row: string[] = [],
    cell = '',
    quoted = false,
    afterQuote = false;
  const input = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < input.length; i++) {
    const c = input[i];
    if (quoted) {
      if (c === '"') {
        if (input[i + 1] === '"') {
          cell += '"';
          i++;
        } else {
          quoted = false;
          afterQuote = true;
        }
      } else cell += c;
    } else if (c === '"' && !cell && !afterQuote) quoted = true;
    else if (c === separator) {
      row.push(cell);
      cell = '';
      afterQuote = false;
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && input[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      cell = '';
      row = [];
      afterQuote = false;
    } else {
      if (afterQuote) throw new Error('Unexpected character after quoted field');
      cell += c;
    }
    if (cell.length > 32768 || rows.length > 100000 || row.length > 1000)
      throw new Error('Delimited data exceeds limits');
  }
  if (quoted) throw new Error('Unterminated quoted field');
  if (cell || row.length || !rows.length) {
    row.push(cell);
    rows.push(row);
  }
  const width = rows.reduce((max, r) => Math.max(max, r.length), 0);
  if (rows.length * width > 200000) throw new Error('Too many cells');
  return rows.map((r) => [...r, ...Array(width - r.length).fill('')]);
}
export function exportRange(
  snapshot: WorkbookSnapshot,
  config: ExportOptions,
): { text: string; report: string[]; requiredPackages: string[] } {
  if (!['csv', 'tsv', 'markdown', 'html', 'latex'].includes(config.format))
    throw new OpenSheetError('INVALID_ARGUMENT', 'Unknown export format');
  const book = new Workbook(snapshot);
  try {
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
    if (config.format === 'csv' || config.format === 'tsv') {
      const result = stringifyDelimited(
        raw.map((row) => row.map((v) => (v && typeof v === 'object' ? v.error : v))),
        config.format === 'csv' ? ',' : '\t',
        opt.safe !== false,
      );
      if (result.sanitized)
        report.push(
          `${result.sanitized} text cells escaped to prevent spreadsheet formula injection.`,
        );
      if (data.merges.length) report.push('Merged cells are flattened in delimited text.');
      return { text: (opt.bom ? '\uFEFF' : '') + result.text, report, requiredPackages };
    }
    const merged = (r: number, c: number) =>
      data.merges.find(
        (m) => r >= m.startRow && r < m.endRow && c >= m.startColumn && c < m.endColumn,
      );
    if (config.format === 'markdown') {
      if (data.merges.length)
        report.push('Markdown has no merged cells; covered cells remain blank.');
      const escaped = rows.map((row) =>
        row.map((v) => v.replace(/\\/g, '\\\\').replace(/\|/g, '\\|').replace(/\r?\n/g, '<br>')),
      );
      const line = (r: string[]) => '| ' + r.join(' | ') + ' |';
      return {
        text: [
          line(escaped[0]),
          line(escaped[0].map(() => '---')),
          ...escaped.slice(1).map(line),
        ].join('\n'),
        report,
        requiredPackages,
      };
    }
    if (config.format === 'html') {
      const text =
        '<table>\n' +
        rows
          .map(
            (row, i) =>
              '  <tr>' +
              row
                .map((v, j) => {
                  const r = b.startRow + i,
                    c = b.startColumn + j,
                    m = merged(r, c);
                  if (m && (r !== m.startRow || c !== m.startColumn)) return '';
                  const tag = i === 0 && opt.header !== false ? 'th' : 'td';
                  const style = book.getStyle(book.getCell(sheet.id, r, c)?.styleId);
                  const css = [
                    style.bold ? 'font-weight:bold' : '',
                    style.italic ? 'font-style:italic' : '',
                    style.underline ? 'text-decoration:underline' : '',
                    style.color ? `color:${style.color}` : '',
                    style.background ? `background-color:${style.background}` : '',
                    style.align ? `text-align:${style.align}` : '',
                    style.fontSize ? `font-size:${style.fontSize}px` : '',
                    style.border ? 'border:1px solid #888' : '',
                  ]
                    .filter(Boolean)
                    .join(';');
                  return `<${tag}${css ? ` style="${css}"` : ''}${m ? ` rowspan="${Math.min(m.endRow, b.endRow) - r}" colspan="${Math.min(m.endColumn, b.endColumn) - c}"` : ''}>${escapeHTML(v).replace(/\n/g, '<br>')}</${tag}>`;
                })
                .join('') +
              '</tr>',
          )
          .join('\n') +
        '\n</table>';
      return { text, report, requiredPackages };
    }
    if (opt.booktabs !== false) requiredPackages.push('booktabs');
    const lines = rows.map((row, i) => {
      const out: string[] = [];
      for (let j = 0; j < row.length; j++) {
        const r = b.startRow + i,
          c = b.startColumn + j,
          m = merged(r, c);
        if (m && c !== m.startColumn) continue;
        let content = escapeLatex(row[j]);
        if (m) {
          const width = Math.min(m.endColumn, b.endColumn) - c,
            height = Math.min(m.endRow, b.endRow) - m.startRow;
          if (r === m.startRow && height > 1) {
            content = `\\multirow{${height}}{*}{${content}}`;
            if (!requiredPackages.includes('multirow')) requiredPackages.push('multirow');
          } else if (r !== m.startRow) content = '';
          if (width > 1) content = `\\multicolumn{${width}}{c}{${content}}`;
        }
        out.push(content);
      }
      return out.join(' & ') + ' \\\\';
    });
    const aligns = rows[0]
      .map((_, j) => {
        const cell = book.getCell(sheet.id, b.startRow, b.startColumn + j);
        const a = book.getStyle(cell?.styleId).align;
        return a === 'right' ? 'r' : a === 'center' ? 'c' : 'l';
      })
      .join('');
    if (opt.booktabs !== false && opt.header !== false && lines.length > 1)
      lines.splice(1, 0, '\\midrule');
    return {
      text: [
        ...requiredPackages.map((p) => `% Requires \\usepackage{${p}}`),
        `\\begin{tabular}{${aligns}}`,
        opt.booktabs !== false ? '\\toprule' : '\\hline',
        ...lines,
        opt.booktabs !== false ? '\\bottomrule' : '\\hline',
        '\\end{tabular}',
      ].join('\n'),
      report,
      requiredPackages,
    };
  } finally {
    book.dispose();
  }
}
