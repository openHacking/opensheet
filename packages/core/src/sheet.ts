import { parseRange } from './address.js';
import { checkBounds } from './model.js';
import { Range } from './range.js';
import { assert, uid, type Rect, type SheetSnapshot } from './types.js';
import type { Workbook } from './workbook.js';

export class Sheet {
  constructor(
    public readonly workbook: Workbook,
    public readonly id: string,
  ) {}
  get name() {
    return this.workbook.sheetData(this.id).name;
  }
  get rowCount() {
    return this.workbook.sheetData(this.id).rowOrder.length;
  }
  get columnCount() {
    return this.workbook.sheetData(this.id).columnOrder.length;
  }
  range(input: string | Rect) {
    const r = parseRange(input);
    checkBounds(this.workbook.sheetData(this.id) as SheetSnapshot, r);
    return new Range(this, r);
  }
  getUsedRange() {
    return this.range(this.workbook.usedRange(this.id));
  }
  insertRows(index: number, count = 1) {
    return this.axis('row', index, count, true);
  }
  deleteRows(index: number, count = 1) {
    return this.axis('row', index, count, false);
  }
  insertColumns(index: number, count = 1) {
    return this.axis('column', index, count, true);
  }
  deleteColumns(index: number, count = 1) {
    return this.axis('column', index, count, false);
  }
  private axis(axis: string, index: number, count: number, insert: boolean) {
    assert(
      Number.isInteger(count) && count > 0 && count <= 100000,
      'INVALID_ARGUMENT',
      'Invalid count',
    );
    return this.workbook.execute({
      type: insert ? 'core.axis.insert' : 'core.axis.delete',
      payload: {
        sheetId: this.id,
        axis,
        index,
        ...(insert ? { ids: Array.from({ length: count }, () => uid(axis)) } : { count }),
      },
    });
  }
  setRowHeight(index: number, size: number) {
    return this.meta('row', index, { size });
  }
  setColumnWidth(index: number, size: number) {
    return this.meta('column', index, { size });
  }
  setRowHidden(index: number, hidden: boolean) {
    return this.meta('row', index, { hidden });
  }
  setColumnHidden(index: number, hidden: boolean) {
    return this.meta('column', index, { hidden });
  }
  private meta(axis: string, index: number, props: object) {
    return this.workbook.execute({
      type: 'core.axis.meta',
      payload: { sheetId: this.id, axis, index, ...props },
    });
  }
  setFreeze(rows: number, columns = 0) {
    return this.workbook.execute({
      type: 'core.sheet.freeze',
      payload: { sheetId: this.id, rows, columns },
    });
  }
  setFilter(filter: { column: number; query: string } | null) {
    return this.workbook.execute({
      type: 'core.sheet.filter',
      payload: { sheetId: this.id, filter },
    });
  }
  sort(range: string | Rect, column: number, direction: 'asc' | 'desc' = 'asc') {
    return this.workbook.execute({
      type: 'core.sheet.sort',
      payload: { sheetId: this.id, range: parseRange(range), column, direction },
    });
  }
}
