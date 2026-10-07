import { escapeLatex } from '../escape.js';
import type { PreparedRange } from '../prepare.js';
import type { ExportResult } from '../types.js';
export function renderLatex(prepared: PreparedRange): ExportResult {
  const { book, sheet, rows, b, data, report, requiredPackages, opt } = prepared;
  const merged = (r: number, c: number) =>
    data.merges.find(
      (m) => r >= m.startRow && r < m.endRow && c >= m.startColumn && c < m.endColumn,
    );
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
}
