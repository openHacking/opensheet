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
  private async read<T>(
    fn: (data: Awaited<ReturnType<typeof this.sheet.workbook.readRange>>, k: string) => T,
    values = true,
  ) {
    const data = await this.sheet.workbook.readRange(this.sheet.id, this.bounds, values);
    const sh = this.sheet.workbook.sheetData(this.sheet.id);
    return this.matrix((r, c) => fn(data, `${sh.rowOrder[r]}:${sh.columnOrder[c]}`));
  }
  getValues() {
    return this.read((data, k) => data.calculated[k] ?? null);
  }
  getDisplayValues() {
    return this.read((data, k) => data.display[k] ?? '');
  }
  getFormulas() {
    return this.read((data, k) => {
      const input = data.cells[k]?.input;
      return input?.type === 'formula' ? input.expression : null;
    }, false);
  }
  getInputs() {
    return this.read((data, k) => data.cells[k]?.input ?? { type: 'blank' as const }, false);
  }
  stream(options: { signal?: AbortSignal; values?: boolean } = {}) {
    return this.sheet.workbook.streamRange(this.sheet.id, this.bounds, options);
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
