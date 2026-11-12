import { evaluateFormula, parseFormula, type Reference } from '@opensheetjs/formula';
import { key, scalarValue, type CellValue, type SheetSnapshot } from './types.js';

export class FormulaCalculation {
  private cache = new Map<string, CellValue>();
  constructor(
    private readSheet: (id: string) => Readonly<SheetSnapshot>,
    private findSheet: (name: string) => string | undefined,
    private dateSystem: '1900' | '1904' = '1900',
  ) {}
  clear() {
    this.cache.clear();
  }
  value(sheetId: string, row: number, column: number, stack = new Set<string>()): CellValue {
    const sh = this.readSheet(sheetId);
    if (row >= sh.rowOrder.length || column >= sh.columnOrder.length || row < 0 || column < 0)
      return { error: '#REF!' };
    const k = `${sheetId}:${row}:${column}`;
    if (this.cache.has(k)) return this.cache.get(k)!;
    if (stack.has(k)) return { error: '#CYCLE!' };
    const c = sh.cells[key(sh.rowOrder[row], sh.columnOrder[column])];
    if (!c) return null;
    if (c.input.type !== 'formula') return scalarValue(c.input);
    if (stack.size >= 128) return { error: '#LIMIT!' };
    stack.add(k);
    let result: CellValue;
    try {
      result = evaluateFormula(
        parseFormula(c.input.expression),
        (ref: Reference) => {
          const target = ref.sheet ? this.findSheet(ref.sheet) : sheetId;
          return target ? this.value(target, ref.row, ref.column, stack) : { error: '#REF!' };
        },
        undefined,
        this.dateSystem,
      );
    } catch (e) {
      result = { error: e instanceof Error ? e.message : '#VALUE!' };
    } finally {
      stack.delete(k);
    }
    if (!(result && typeof result === 'object' && ['#LIMIT!', '#CYCLE!'].includes(result.error)))
      this.cache.set(k, result);
    return result;
  }
}
