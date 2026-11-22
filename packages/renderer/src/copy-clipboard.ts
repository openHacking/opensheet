import { OpenSheetError, toInput, type Range, type CellStyle } from '@opensheetjs/core';
import { encodeClipboardText } from './clipboard-text.js';
import { copyText } from './clipboard.js';
import { clipboardBatches } from './clipboard-batches.js';

function escape(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
function css(style: CellStyle) {
  return Object.entries({
    'font-weight': style.bold === undefined ? undefined : style.bold ? 'bold' : 'normal',
    'font-style': style.italic === undefined ? undefined : style.italic ? 'italic' : 'normal',
    'text-decoration': style.underline ? 'underline' : undefined,
    color: style.color,
    'background-color': style.background,
    'text-align': style.align,
    'font-size': style.fontSize === undefined ? undefined : `${style.fontSize}px`,
    'white-space': style.wrap ? 'pre-wrap' : 'pre',
  })
    .filter(([, value]) => value !== undefined)
    .map(([key, value]) => `${key}:${value}`)
    .join(';');
}

export async function copyClipboard(range: Range): Promise<{ text: string; html: string }> {
  const book = range.sheet.workbook,
    bounds = { ...range.bounds };
  const revision = book.revision;
  if (
    book
      .sheetData(range.sheet.id)
      .merges.some(
        (m) =>
          m.startRow < bounds.endRow &&
          m.endRow > bounds.startRow &&
          m.startColumn < bounds.endColumn &&
          m.endColumn > bounds.startColumn,
      )
  )
    throw new OpenSheetError('INVALID_RANGE', 'Unmerge before copying structured cells');
  const rows: string[][] = [];
  let html = `<table data-opensheet-clipboard="1" data-row="${bounds.startRow}" data-column="${bounds.startColumn}">`;
  for (const batch of clipboardBatches(book, range.sheet.id, bounds)) {
    const data = await book.readCells(range.sheet.id, batch, {
      fields: ['input', 'value', 'display', 'style', 'numberFormat', 'note', 'link'],
    });
    if (data.revision !== revision || book.revision !== revision)
      throw new OpenSheetError('REVISION_CONFLICT', 'Workbook changed during copy');
    for (const cell of data.cells) {
      if (cell.column === bounds.startColumn) {
        html += '<tr>';
        rows.push([]);
      }
      const value = cell.value ?? null;
      if (value && typeof value === 'object' && value.error === '#LIMIT!')
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Cannot copy an incomplete calculation');
      const scalar =
        value && typeof value === 'object' ? { type: 'error', code: value.error } : toInput(value);
      const style = {
        ...cell.style,
        ...(cell.numberFormat !== undefined ? { numberFormat: cell.numberFormat } : {}),
      };
      const payload = {
        input: cell.input ?? { type: 'blank' },
        value: scalar,
        style,
        note: cell.note ?? null,
        link: cell.link ?? null,
      };
      const display = cell.display ?? '';
      rows[rows.length - 1].push(display);
      html += `<td data-opensheet-cell="${escape(JSON.stringify(payload))}" style="${escape(css(style))}">${escape(display).replace(/\r\n|\r|\n/g, '<br>')}</td>`;
      if (cell.column === bounds.endColumn - 1) html += '</tr>';
      if (html.length + 8 > 2 * 1024 * 1024)
        throw new OpenSheetError('LIMIT_EXCEEDED', 'Structured clipboard is too large');
    }
  }
  return { text: encodeClipboardText(rows), html: html + '</table>' };
}

export async function writeClipboard(range: Range) {
  const clipboard = navigator.clipboard;
  if (!clipboard) throw new Error('Clipboard is unavailable. Use Ctrl+C or Command+C.');
  if (typeof ClipboardItem === 'undefined' || !clipboard.write) {
    await clipboard.writeText(await copyText(range));
    return;
  }
  const data = copyClipboard(range);
  const text = data.then(({ text }) => new Blob([text], { type: 'text/plain' }));
  const html = data.then(({ html }) => new Blob([html], { type: 'text/html' }));
  // Observe both producers immediately, including after a rejected clipboard write.
  // A failed permission request must not wait for a stalled workbook read.
  const results = Promise.allSettled([text, html]);
  await clipboard.write([new ClipboardItem({ 'text/plain': text, 'text/html': html })]);
  for (const result of await results) if (result.status === 'rejected') throw result.reason;
}
