import { parseRange } from './address.js';
import { FormulaCalculation } from './calculation.js';
import { formatValue } from './format.js';
import { checkBounds } from './model.js';
import { key, type Rect, type WorkbookSnapshot } from './types.js';
/** Explicit materialized snapshot reader for file adapters; never backs an editor. */
export class SnapshotReader {
  private calculator: FormulaCalculation;
  constructor(private snapshot: WorkbookSnapshot) {
    this.calculator = new FormulaCalculation(
      (id) => this.sheetData(id),
      (name) => snapshot.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase())?.id,
    );
  }
  sheetData(id: string) {
    const sh = this.snapshot.sheets.find((s) => s.id === id);
    if (!sh) throw new Error('Unknown sheet');
    return sh;
  }
  getStyle(id?: string) {
    return this.snapshot.styles[id ?? ''] ?? {};
  }
  getCell(id: string, r: number, c: number) {
    const sh = this.sheetData(id);
    return sh.cells[key(sh.rowOrder[r], sh.columnOrder[c])];
  }
  value(id: string, r: number, c: number) {
    return this.calculator.value(id, r, c);
  }
  display(id: string, r: number, c: number) {
    const cell = this.getCell(id, r, c);
    return formatValue(
      this.value(id, r, c),
      cell?.numberFormat ?? this.getStyle(cell?.styleId).numberFormat,
      this.snapshot.dateSystem,
    );
  }
  getSheetById(id: string) {
    if (!this.snapshot.sheets.some((s) => s.id === id)) return;
    const data = this.sheetData(id);
    const range = (input: string | Rect) => {
      const bounds = parseRange(input);
      checkBounds(data, bounds);
      const matrix = <T>(fn: (r: number, c: number) => T) =>
        Array.from({ length: bounds.endRow - bounds.startRow }, (_, i) =>
          Array.from({ length: bounds.endColumn - bounds.startColumn }, (_, j) =>
            fn(bounds.startRow + i, bounds.startColumn + j),
          ),
        );
      return {
        bounds,
        getValues: () => matrix((r, c) => this.value(id, r, c)),
        getDisplayValues: () => matrix((r, c) => this.display(id, r, c)),
      };
    };
    return {
      id,
      range,
      getUsedRange: () => {
        const rows = new Map(data.rowOrder.map((k, i) => [k, i])),
          cols = new Map(data.columnOrder.map((k, i) => [k, i]));
        let endRow = 1,
          endColumn = 1;
        for (const cell of Object.values(data.cells)) {
          endRow = Math.max(endRow, rows.get(cell.rowId)! + 1);
          endColumn = Math.max(endColumn, cols.get(cell.columnId)! + 1);
        }
        for (const m of data.merges) {
          endRow = Math.max(endRow, m.endRow);
          endColumn = Math.max(endColumn, m.endColumn);
        }
        return range({ startRow: 0, startColumn: 0, endRow, endColumn });
      },
    };
  }
  dispose() {
    this.calculator.clear();
  }
}
