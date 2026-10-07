import { parseRange, rangeAddress } from './address.js';
import { parseInput } from './input.js';
import type { Sheet } from './sheet.js';
import {
  assert,
  toInput,
  type CellInput,
  type CellStyle,
  type Primitive,
  type Rect,
} from './types.js';

export class Range {
  constructor(
    public readonly sheet: Sheet,
    public readonly bounds: Rect,
  ) {}
  get address() {
    return rangeAddress(this.bounds);
  }
  private matrix<T>(fn: (r: number, c: number) => T): T[][] {
    const b = this.bounds;
    return Array.from({ length: b.endRow - b.startRow }, (_, i) =>
      Array.from({ length: b.endColumn - b.startColumn }, (_, j) =>
        fn(b.startRow + i, b.startColumn + j),
      ),
    );
  }
  getValues() {
    return this.matrix((r, c) => this.sheet.workbook.value(this.sheet.id, r, c));
  }
  getDisplayValues() {
    return this.matrix((r, c) => this.sheet.workbook.display(this.sheet.id, r, c));
  }
  getFormulas() {
    return this.matrix((r, c) => {
      const input = this.sheet.workbook.getCell(this.sheet.id, r, c)?.input;
      return input?.type === 'formula' ? input.expression : null;
    });
  }
  getInputs() {
    return this.matrix(
      (r, c) => this.sheet.workbook.getCell(this.sheet.id, r, c)?.input ?? { type: 'blank' },
    );
  }
  private write<T>(matrix: T[][], map: (value: T) => CellInput) {
    const b = this.bounds;
    assert(
      matrix.length === b.endRow - b.startRow &&
        matrix.every((row) => row.length === b.endColumn - b.startColumn),
      'INVALID_ARGUMENT',
      'Matrix dimensions must match the range',
    );
    return this.sheet.workbook.execute({
      type: 'core.cells.set',
      payload: {
        sheetId: this.sheet.id,
        cells: matrix.flatMap((row, i) =>
          row.map((value, j) => ({
            row: b.startRow + i,
            column: b.startColumn + j,
            input: map(value),
          })),
        ),
      },
    });
  }
  setValues(values: Primitive[][]) {
    return this.write(values, toInput);
  }
  copyTo(target: string | Rect) {
    return this.command('core.cells.copy', { target: parseRange(target) });
  }
  fillDown() {
    return this.command('core.cells.fillDown');
  }
  setFormulas(values: string[][]) {
    return this.write(values, (expression) => ({
      type: 'formula',
      expression: expression.replace(/^=/, ''),
    }));
  }
  setInput(value: string) {
    assert(
      this.bounds.endRow - this.bounds.startRow === 1 &&
        this.bounds.endColumn - this.bounds.startColumn === 1,
      'INVALID_RANGE',
      'setInput expects one cell',
    );
    return this.write([[value]], parseInput);
  }
  setStyle(style: CellStyle) {
    return this.command('core.cells.style', { style });
  }
  clear(options: { all?: boolean } = {}) {
    return this.command('core.cells.clear', { all: options.all ?? false });
  }
  merge(options: { discardCoveredValues?: boolean } = {}) {
    return this.command('core.cells.merge', {
      discardCoveredValues: options.discardCoveredValues ?? false,
    });
  }
  unmerge() {
    return this.command('core.cells.unmerge');
  }
  private command(type: string, extra: object = {}) {
    return this.sheet.workbook.execute({
      type,
      payload: { sheetId: this.sheet.id, range: this.bounds, ...extra },
    });
  }
}
