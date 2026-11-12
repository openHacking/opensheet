import { parseRange, rangeAddress } from './address.js';
import { selectiveCopyCell } from './selective-copy.js';
import { parseInput } from './input.js';
import type { Sheet } from './sheet.js';
import {
  assert,
  toInput,
  type CellInput,
  type CellLink,
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
  private matrix<T>(fn: (r: number, c: number) => T, b = this.bounds): T[][] {
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
    const book = this.sheet.workbook;
    const revision = book.revision;
    const sh = book.sheetData(this.sheet.id);
    const bounds = { ...this.bounds };
    const data = await book.readRange(this.sheet.id, bounds, values);
    assert(data.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during read');
    return this.matrix((r, c) => fn(data, `${sh.rowOrder[r]}:${sh.columnOrder[c]}`), bounds);
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
  getNotes() {
    return this.read((data, k) => data.cells[k]?.note ?? null, false);
  }
  getLinks() {
    return this.read((data, k) => data.cells[k]?.link ?? null, false);
  }
  setNote(note: string | null) {
    return this.command('core.cells.annotation', { note });
  }
  setLink(link: CellLink | null) {
    return this.command('core.cells.annotation', { link });
  }
  getInputs() {
    return this.read((data, k) => data.cells[k]?.input ?? { type: 'blank' as const }, false);
  }
  stream(options: { signal?: AbortSignal; values?: boolean } = {}) {
    return this.sheet.workbook.streamRange(this.sheet.id, this.bounds, options);
  }
  private write<T>(matrix: T[][], map: (value: T) => CellInput) {
    const b = { ...this.bounds };
    const height = b.endRow - b.startRow;
    const width = b.endColumn - b.startColumn;
    assert(
      Array.isArray(matrix) && matrix.length === height,
      'INVALID_ARGUMENT',
      'Matrix dimensions must match the range',
    );
    const cells: Array<{ row: number; column: number; input: CellInput }> = [];
    for (let row = 0; row < height; row++) {
      assert(
        Object.hasOwn(matrix, row) && Array.isArray(matrix[row]) && matrix[row].length === width,
        'INVALID_ARGUMENT',
        'Matrix dimensions must match the range',
      );
      for (let column = 0; column < width; column++) {
        assert(Object.hasOwn(matrix[row], column), 'INVALID_ARGUMENT', 'Matrix must be dense');
        cells.push({
          row: b.startRow + row,
          column: b.startColumn + column,
          input: map(matrix[row][column]),
        });
      }
    }
    return this.sheet.workbook.execute({
      type: 'core.cells.set',
      payload: { sheetId: this.sheet.id, cells },
    });
  }
  setValues(values: Primitive[][]) {
    return this.write(values, toInput);
  }
  copyTo(
    target: string | Rect,
    options: { mode?: 'all' | 'values' | 'formulas' | 'formats' } = {},
  ) {
    if (options.mode !== undefined && options.mode !== 'all')
      return this.copySpecial(target, options.mode);
    return this.command('core.cells.copy', { target: parseRange(target) });
  }
  private async copySpecial(target: string | Rect, mode: 'values' | 'formulas' | 'formats') {
    assert(
      ['values', 'formulas', 'formats'].includes(mode),
      'INVALID_ARGUMENT',
      'Unknown copy mode',
    );
    const destination = parseRange(target),
      source = { ...this.bounds },
      book = this.sheet.workbook;
    if (book.isTransactionScope)
      return book.execute({
        type: 'core.cells.copySpecial',
        payload: { sheetId: this.sheet.id, range: source, target: destination, mode },
      });
    const revision = book.revision;
    assert(
      destination.endRow - destination.startRow === source.endRow - source.startRow &&
        destination.endColumn - destination.startColumn === source.endColumn - source.startColumn,
      'INVALID_RANGE',
      'Copy ranges must have equal dimensions',
    );
    const sh = book.sheetData(this.sheet.id);
    assert(
      !sh.merges.some(
        (m) =>
          m.startRow < source.endRow &&
          m.endRow > source.startRow &&
          m.startColumn < source.endColumn &&
          m.endColumn > source.startColumn,
      ),
      'INVALID_RANGE',
      'Unmerge before copying',
    );
    const data = await book.readCells(this.sheet.id, source, {
      fields:
        mode === 'values' ? ['value'] : mode === 'formulas' ? ['input'] : ['style', 'numberFormat'],
    });
    assert(data.revision === revision, 'REVISION_CONFLICT', 'Workbook changed during copy');
    const dr = destination.startRow - source.startRow,
      dc = destination.startColumn - source.startColumn;
    const cells = data.cells.map((cell) => selectiveCopyCell(cell, mode, dr, dc));
    return book.execute({
      protocolVersion: 1,
      commandId: crypto.randomUUID(),
      workbookId: book.id,
      baseRevision: data.revision,
      type: 'core.cells.paste',
      payload: { sheetId: this.sheet.id, cells },
    });
  }
  fillSeries(options: {
    start: number;
    step?: number;
    axis?: 'row' | 'column';
    calendar?: 'month' | 'year';
  }) {
    return this.command('core.cells.fillSeries', options);
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
