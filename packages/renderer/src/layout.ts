import type { Workbook } from '@opensheetjs/core';
import { LEFT, TOP } from './geometry.js';
export class GridLayout {
  rowOffsets: number[] = [];
  columnOffsets: number[] = [];
  dirty = true;
  rebuild(book: Workbook, sheetId: string, zoom: number, spacer: HTMLElement) {
    const sh = book.sheetData(sheetId);
    const filteredRows = new Set(book.filteredRows(sheetId));
    this.rowOffsets = [0];
    this.columnOffsets = [0];
    for (let i = 0; i < sh.rowOrder.length; i++) {
      const m = sh.rows[sh.rowOrder[i]];
      const filtered = filteredRows.has(i);
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
  /** Visible nonzero-sized entries, including frozen and scrolling panes. */
  visiblePositions(offsets: number[], scroll: number, extent: number, frozen: number) {
    const out: number[] = [];
    const end = offsets.length - 1;
    const collect = (start: number, limit: number, shift: number) => {
      for (let i = start; i < limit && offsets[i] - shift < extent; ) {
        if (offsets[i + 1] > offsets[i]) {
          out.push(i++);
        } else if (i + 1 >= limit || offsets[i + 2] > offsets[i]) {
          i++;
        } else {
          // Short runs are cheaper to scan; bound that scan before searching a long run.
          const scanEnd = Math.min(i + 8, limit);
          while (i < scanEnd && offsets[i + 1] === offsets[i]) i++;
          if (i === scanEnd && i < limit && offsets[i + 1] === offsets[i])
            i = Math.max(i + 1, this.index(offsets, offsets[i]));
        }
      }
    };
    collect(0, Math.min(frozen, end), 0);
    collect(Math.max(frozen, this.index(offsets, scroll + offsets[frozen])), end, scroll);
    return out;
  }
  /** Resolve a hidden position in the requested direction, falling back at an edge. */
  visibleIndex(offsets: number[], index: number, direction: number) {
    if (!direction || offsets[index] !== offsets[index + 1]) return index;
    const target = offsets[index];
    let lo = 0,
      hi = offsets.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (offsets[mid] < target) lo = mid + 1;
      else hi = mid;
    }
    const before = lo - 1;
    hi = offsets.length;
    while (lo < hi) {
      const mid = (lo + hi) >>> 1;
      if (offsets[mid] <= target) lo = mid + 1;
      else hi = mid;
    }
    const after = lo - 1;
    const valid = (value: number) => value >= 0 && value < offsets.length - 1;
    const first = direction < 0 ? before : after;
    const fallback = direction < 0 ? after : before;
    return valid(first) ? first : valid(fallback) ? fallback : index;
  }
  hit(
    e: MouseEvent,
    element: HTMLElement,
    freeze: { rows: number; columns: number },
    scroll: { scrollLeft: number; scrollTop: number; clientWidth: number; clientHeight: number },
  ) {
    const rect = element.getBoundingClientRect(),
      x = e.clientX - rect.left - LEFT,
      y = e.clientY - rect.top - TOP;
    if (
      x < 0 ||
      y < 0 ||
      e.clientX >= rect.left + scroll.clientWidth ||
      e.clientY >= rect.top + scroll.clientHeight
    )
      return null;
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
