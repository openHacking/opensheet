import type { WorksheetFilter } from './filter-schema.js';
import type { SortColumn } from './sort-values.js';
import { parseRange } from './address.js';
import { checkBounds, checkDimensions } from './model.js';
import { Range } from './range.js';
import { assert, type Rect } from './types.js';
import type { Workbook } from './workbook.js';

export class Sheet {
  constructor(
    public readonly workbook: Workbook,
    public readonly id: string,
  ) {}
  get name() {
    return this.workbook.sheetCommandName(this.id);
  }
  get rowCount() {
    return this.workbook.sheetCommandState(this.id).rows;
  }
  get columnCount() {
    return this.workbook.sheetCommandState(this.id).columns;
  }
  range(input: string | Rect) {
    const r = parseRange(input);
    const dimensions = this.workbook.sheetCommandState(this.id);
    checkDimensions(dimensions.rows, dimensions.columns, r);
    return new Range(this, r);
  }
  getUsedRange() {
    const used = this.workbook.usedRange(this.id);
    assert(used, 'INVALID_ARGUMENT', 'Sheet has no committed range');
    const range = parseRange(used);
    checkBounds(this.workbook.sheetData(this.id), range);
    return new Range(this, range);
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
    const sheet = this.workbook.sheetCommandState(this.id);
    const next = axis === 'row' ? sheet.nextRowId : sheet.nextColumnId;
    return this.workbook.execute({
      type: insert ? 'core.axis.insert' : 'core.axis.delete',
      payload: {
        sheetId: this.id,
        axis,
        index,
        ...(insert
          ? {
              ids: Array.from({ length: count }, (_, i) => next + i),
            }
          : { count }),
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
  setRowsHeight(index: number, count: number, size: number) {
    return this.resize('row', index, count, size);
  }
  setColumnsWidth(index: number, count: number, size: number) {
    return this.resize('column', index, count, size);
  }
  private resize(axis: 'row' | 'column', index: number, count: number, size: number) {
    return this.workbook.execute({
      type: 'core.axis.resize',
      payload: { sheetId: this.id, axis, index, count, size },
    });
  }
  setRowsHidden(index: number, count: number, hidden: boolean) {
    return this.visibility('row', index, count, hidden);
  }
  setColumnsHidden(index: number, count: number, hidden: boolean) {
    return this.visibility('column', index, count, hidden);
  }
  setAllRowsHidden(hidden: boolean) {
    return this.workbook.execute({
      type: 'core.axis.visibilityAll',
      payload: { sheetId: this.id, axis: 'row', hidden },
    });
  }
  setAllColumnsHidden(hidden: boolean) {
    return this.workbook.execute({
      type: 'core.axis.visibilityAll',
      payload: { sheetId: this.id, axis: 'column', hidden },
    });
  }
  private visibility(axis: 'row' | 'column', index: number, count: number, hidden: boolean) {
    return this.workbook.execute({
      type: 'core.axis.visibility',
      payload: { sheetId: this.id, axis, index, count, hidden },
    });
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
  setFilter(filter: WorksheetFilter) {
    return this.workbook.execute({
      type: 'core.sheet.filter',
      payload: { sheetId: this.id, filter },
    });
  }
  sort(
    range: string | Rect,
    column: number | readonly SortColumn[],
    direction: 'asc' | 'desc' = 'asc',
  ) {
    return this.workbook.execute({
      type: 'core.sheet.sort',
      payload: {
        sheetId: this.id,
        range: parseRange(range),
        ...(typeof column === 'number'
          ? { column, direction }
          : { keys: column.map((key) => ({ ...key })) }),
      },
    });
  }
}
