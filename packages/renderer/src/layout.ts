import type { Workbook } from '@opensheetjs/core';
import { LEFT, TOP } from './geometry.js';
export class GridLayout {
  rowOffsets: number[] = [];
  columnOffsets: number[] = [];
  dirty = true;
  rebuild(book: Workbook, sheetId: string, zoom: number, spacer: HTMLElement) {
    const sh = book.sheetData(sheetId);
    this.rowOffsets = [0];
    this.columnOffsets = [0];
    for (let i = 0; i < sh.rowOrder.length; i++) {
      const m = sh.rows[sh.rowOrder[i]];
      const filtered =
        sh.filter &&
        i > 0 &&
        !book
          .display(sheetId, i, sh.filter.column)
          .toLowerCase()
          .includes(sh.filter.query.toLowerCase());
      this.rowOffsets.push(
        this.rowOffsets[i] + (m?.hidden || filtered ? 0 : (m?.size ?? 30) * zoom),
      );
    }
    for (let i = 0; i < sh.columnOrder.length; i++) {
      const m = sh.columns[sh.columnOrder[i]];
      this.columnOffsets.push(this.columnOffsets[i] + (m?.hidden ? 0 : (m?.size ?? 128) * zoom));
    }
    spacer.style.width = `${LEFT + this.columnOffsets.at(-1)!}px`;
    spacer.style.height = `${TOP + this.rowOffsets.at(-1)!}px`;
    this.dirty = false;
  }
  position(
    row: number,
    column: number,
    freeze: { rows: number; columns: number },
    scroll: { scrollLeft: number; scrollTop: number },
  ) {
    return {
      x: LEFT + this.columnOffsets[column] - (column >= freeze.columns ? scroll.scrollLeft : 0),
      y: TOP + this.rowOffsets[row] - (row >= freeze.rows ? scroll.scrollTop : 0),
    };
  }
  index(offsets: number[], value: number) {
    let lo = 0,
      hi = offsets.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (offsets[mid] <= value) lo = mid;
      else hi = mid - 1;
    }
    return Math.min(lo, offsets.length - 2);
  }
  hit(
    e: PointerEvent,
    element: HTMLElement,
    freeze: { rows: number; columns: number },
    scroll: { scrollLeft: number; scrollTop: number },
  ) {
    const rect = element.getBoundingClientRect(),
      x = e.clientX - rect.left - LEFT,
      y = e.clientY - rect.top - TOP;
    if (x < 0 || y < 0 || e.clientX > rect.right || e.clientY > rect.bottom) return null;
    const f = freeze;
    return {
      row: this.index(this.rowOffsets, y + (y >= this.rowOffsets[f.rows] ? scroll.scrollTop : 0)),
      column: this.index(
        this.columnOffsets,
        x + (x >= this.columnOffsets[f.columns] ? scroll.scrollLeft : 0),
      ),
    };
  }
}
