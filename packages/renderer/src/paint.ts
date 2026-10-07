import { address, columnName, type Selection, type Workbook } from '@opensheetjs/core';
import { LEFT, TOP } from './geometry.js';
import type { GridLayout } from './layout.js';
import type { GridSelection } from './selection.js';
export function paintGrid(
  book: Workbook,
  sheetId: string,
  selected: Selection,
  zoom: number,
  layout: GridLayout,
  selection: GridSelection,
  element: HTMLElement,
  canvas: HTMLCanvasElement,
  ariaCell: HTMLElement,
  scroll: HTMLElement,
) {
  const position = (row: number, column: number) =>
    layout.position(row, column, book.sheetData(sheetId).freeze, scroll);
  const w = element.clientWidth,
    h = element.clientHeight,
    dpr = window.devicePixelRatio || 1;
  if (!w || !h) return;
  const pixelWidth = Math.round(w * dpr),
    pixelHeight = Math.round(h * dpr);
  if (canvas.width !== pixelWidth || canvas.height !== pixelHeight) {
    canvas.width = pixelWidth;
    canvas.height = pixelHeight;
    canvas.style.width = `${w}px`;
    canvas.style.height = `${h}px`;
  }
  const ctx = canvas.getContext('2d')!;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, w, h);
  const sh = book.sheetData(sheetId),
    f = sh.freeze;
  const freezeX = LEFT + layout.columnOffsets[f.columns],
    freezeY = TOP + layout.rowOffsets[f.rows];
  const visible = (offsets: number[], scroll: number, extent: number, frozen: number) => {
    const out: number[] = [];
    for (let i = 0; i < Math.min(frozen, offsets.length - 1) && offsets[i] < extent; i++)
      if (offsets[i + 1] > offsets[i]) out.push(i);
    const start = Math.max(frozen, layout.index(offsets, scroll + offsets[frozen]));
    for (let i = start; i < offsets.length - 1 && offsets[i] - scroll < extent; i++)
      if (offsets[i + 1] > offsets[i]) out.push(i);
    return out;
  };
  const rows = visible(layout.rowOffsets, scroll.scrollTop, h - TOP, f.rows),
    cols = visible(layout.columnOffsets, scroll.scrollLeft, w - LEFT, f.columns);
  const drawCell = (r: number, c: number) => {
    const merge = selection.mergeAt(r, c);
    if (merge && (r !== merge.startRow || c !== merge.startColumn)) return;
    const endColumn = merge?.endColumn ?? c + 1,
      endRow = merge?.endRow ?? r + 1,
      cw = layout.columnOffsets[endColumn] - layout.columnOffsets[c],
      rh = layout.rowOffsets[merge?.endRow ?? r + 1] - layout.rowOffsets[r];
    const cell = book.peekCell(sheetId, r, c),
      style = book.getStyle(cell?.styleId);
    const horizontal = c < f.columns && endColumn > f.columns ? [false, true] : [c >= f.columns];
    const vertical = r < f.rows && endRow > f.rows ? [false, true] : [r >= f.rows];
    for (const scrollX of horizontal)
      for (const scrollY of vertical) {
        const pos = {
          x: LEFT + layout.columnOffsets[c] - (scrollX ? scroll.scrollLeft : 0),
          y: TOP + layout.rowOffsets[r] - (scrollY ? scroll.scrollTop : 0),
        };
        const clipX = scrollX ? freezeX : LEFT,
          clipY = scrollY ? freezeY : TOP,
          clipRight = scrollX ? w : freezeX,
          clipBottom = scrollY ? h : freezeY;
        if (clipRight <= clipX || clipBottom <= clipY) continue;
        ctx.save();
        ctx.beginPath();
        ctx.rect(clipX, clipY, clipRight - clipX, clipBottom - clipY);
        ctx.clip();
        ctx.fillStyle = style.background ?? '#ffffff';
        ctx.fillRect(pos.x, pos.y, cw, rh);
        if (
          r >= selected.startRow &&
          r < selected.endRow &&
          c >= selected.startColumn &&
          c < selected.endColumn
        ) {
          ctx.fillStyle = 'rgba(16,130,94,.06)';
          ctx.fillRect(pos.x, pos.y, cw, rh);
        }
        ctx.strokeStyle = style.border ? '#8ca098' : '#e9eeeb';
        ctx.lineWidth = 1;
        ctx.strokeRect(pos.x + 0.5, pos.y + 0.5, cw, rh);
        ctx.beginPath();
        ctx.rect(pos.x + 3, pos.y + 1, Math.max(0, cw - 6), Math.max(0, rh - 2));
        ctx.clip();
        ctx.fillStyle = style.color ?? '#293d35';
        ctx.font = `${style.italic ? 'italic ' : ''}${style.bold ? '600 ' : ''}${(style.fontSize ?? 13) * zoom}px Inter, -apple-system, sans-serif`;
        ctx.textBaseline = 'middle';
        const raw = book.peekValue(sheetId, r, c);
        const align = style.align ?? (typeof raw === 'number' ? 'right' : 'left');
        ctx.textAlign = align;
        const x =
          align === 'right' ? pos.x + cw - 10 : align === 'center' ? pos.x + cw / 2 : pos.x + 10;
        const text = book.peekDisplay(sheetId, r, c);
        if (style.wrap) {
          const lineHeight = (style.fontSize ?? 13) * zoom * 1.35;
          const lines: string[] = [];
          let line = '';
          for (const char of text.slice(0, 4096)) {
            if (char === '\n' || ctx.measureText(line + char).width > cw - 20) {
              lines.push(line);
              line = char === '\n' ? '' : char;
            } else line += char;
            if (lines.length > Math.ceil(rh / lineHeight)) break;
          }
          if (line) lines.push(line);
          lines.forEach((part, i) => ctx.fillText(part, x, pos.y + lineHeight * (i + 0.6)));
        } else ctx.fillText(text.replace(/\r?\n/g, ' '), x, pos.y + rh / 2);
        if (style.underline) {
          const length = Math.min(ctx.measureText(text).width, cw - 20);
          ctx.fillRect(
            align === 'right' ? x - length : align === 'center' ? x - length / 2 : x,
            pos.y + rh / 2 + 8,
            length,
            1,
          );
        }
        ctx.restore();
      }
  };
  // Include offscreen merge anchors whose merged rectangle intersects the viewport.
  const anchors = new Set<string>();
  for (const r of rows)
    for (const c of cols) {
      const m = selection.mergeAt(r, c);
      if (m) anchors.add(`${m.startRow},${m.startColumn}`);
      else drawCell(r, c);
    }
  for (const a of anchors) {
    const [r, c] = a.split(',').map(Number);
    drawCell(r, c);
  }
  const s = selected;
  for (const [scrollX, scrollY, x1, y1, x2, y2] of [
    [0, 0, LEFT, TOP, freezeX, freezeY],
    [1, 0, freezeX, TOP, w, freezeY],
    [0, 1, LEFT, freezeY, freezeX, h],
    [1, 1, freezeX, freezeY, w, h],
  ]) {
    if (x2 <= x1 || y2 <= y1) continue;
    const x = LEFT + layout.columnOffsets[s.startColumn] - (scrollX ? scroll.scrollLeft : 0),
      y = TOP + layout.rowOffsets[s.startRow] - (scrollY ? scroll.scrollTop : 0),
      endX = LEFT + layout.columnOffsets[s.endColumn] - (scrollX ? scroll.scrollLeft : 0),
      endY = TOP + layout.rowOffsets[s.endRow] - (scrollY ? scroll.scrollTop : 0);
    ctx.save();
    ctx.beginPath();
    ctx.rect(x1, y1, x2 - x1, y2 - y1);
    ctx.clip();
    ctx.strokeStyle = '#168462';
    ctx.lineWidth = 2;
    ctx.strokeRect(x + 1, y + 1, endX - x - 2, endY - y - 2);
    ctx.restore();
  }
  ctx.fillStyle = '#f6f8f6';
  ctx.fillRect(0, 0, w, TOP);
  ctx.fillRect(0, 0, LEFT, h);
  ctx.font = '11px Inter, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.strokeStyle = '#e3e9e5';
  for (const c of cols) {
    const { x } = position(0, c),
      cw = layout.columnOffsets[c + 1] - layout.columnOffsets[c];
    if (x < LEFT) continue;
    ctx.save();
    ctx.beginPath();
    const headerLeft = c < f.columns ? LEFT : freezeX;
    ctx.rect(headerLeft, 0, w - headerLeft, TOP);
    ctx.clip();
    ctx.fillStyle = c >= s.startColumn && c < s.endColumn ? '#deeee6' : '#f6f8f6';
    ctx.fillRect(x, 0, cw, TOP);
    ctx.fillStyle = '#667b6f';
    ctx.fillText(columnName(c), x + cw / 2, TOP / 2);
    ctx.strokeRect(x + 0.5, 0.5, cw, TOP);
    ctx.restore();
  }
  for (const r of rows) {
    const { y } = position(r, 0),
      rh = layout.rowOffsets[r + 1] - layout.rowOffsets[r];
    if (y < TOP) continue;
    ctx.save();
    ctx.beginPath();
    const headerTop = r < f.rows ? TOP : freezeY;
    ctx.rect(0, headerTop, LEFT, h - headerTop);
    ctx.clip();
    ctx.fillStyle = r >= s.startRow && r < s.endRow ? '#deeee6' : '#f6f8f6';
    ctx.fillRect(0, y, LEFT, rh);
    ctx.fillStyle = '#667b6f';
    ctx.fillText(String(r + 1), LEFT / 2, y + rh / 2);
    ctx.strokeRect(0.5, y + 0.5, LEFT, rh);
    ctx.restore();
  }
  if (f.rows) {
    ctx.strokeStyle = '#a7bdb0';
    ctx.beginPath();
    ctx.moveTo(LEFT, TOP + layout.rowOffsets[f.rows]);
    ctx.lineTo(w, TOP + layout.rowOffsets[f.rows]);
    ctx.stroke();
  }
  if (f.columns) {
    ctx.strokeStyle = '#a7bdb0';
    ctx.beginPath();
    ctx.moveTo(LEFT + layout.columnOffsets[f.columns], TOP);
    ctx.lineTo(LEFT + layout.columnOffsets[f.columns], h);
    ctx.stroke();
  }
  element.setAttribute('aria-rowcount', String(sh.rowOrder.length));
  element.setAttribute('aria-colcount', String(sh.columnOrder.length));
  ariaCell.setAttribute('aria-rowindex', String(s.startRow + 1));
  ariaCell.setAttribute('aria-colindex', String(s.startColumn + 1));
  ariaCell.textContent = `${address(s.startRow, s.startColumn)}: ${book.peekDisplay(sheetId, s.startRow, s.startColumn)}`;
}
