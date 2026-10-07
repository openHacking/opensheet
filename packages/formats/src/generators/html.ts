import { escapeHTML } from '../escape.js';
import type { PreparedRange } from '../prepare.js';
import type { ExportResult } from '../types.js';
export function renderHtml(prepared: PreparedRange): ExportResult {
  const { book, sheet, rows, b, data, report, requiredPackages, opt } = prepared;
  const merged = (r: number, c: number) =>
    data.merges.find(
      (m) => r >= m.startRow && r < m.endRow && c >= m.startColumn && c < m.endColumn,
    );

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
