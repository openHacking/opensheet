import { evaluateFormula, parseFormula, type Reference } from '@opensheetjs/formula';
import { key, scalarValue, type CellValue, type SheetSnapshot } from './types.js';
import { ByteCache } from './storage.js';

export class FormulaCalculation {
  private cache = new ByteCache<{ value: CellValue; depth: number }>(4 * 1024 * 1024);
  get cacheBytes() {
    return this.cache.bytes;
  }
  constructor(
    private readSheet: (id: string) => Readonly<SheetSnapshot>,
    private findSheet: (name: string) => string | undefined,
    private dateSystem: '1900' | '1904' = '1900',
  ) {}
  clear() {
    this.cache.clear();
  }
  value(
    sheetId: string,
    row: number,
    column: number,
    stack = new Set<string>(),
    parent?: { depth: number },
  ): CellValue {
    const sh = this.readSheet(sheetId);
    if (row >= sh.rowOrder.length || column >= sh.columnOrder.length || row < 0 || column < 0)
      return { error: '#REF!' };
    const k = `${sheetId}:${row}:${column}`;
    if (stack.has(k)) return { error: '#CYCLE!' };
    const cached = this.cache.get(k);
    if (cached !== undefined) {
      if (stack.size + cached.depth > 128) return { error: '#LIMIT!' };
      if (parent) parent.depth = Math.max(parent.depth, cached.depth);
      return cached.value;
    }
    const c = sh.cells[key(sh.rowOrder[row], sh.columnOrder[column])];
    if (!c) return null;
    if (c.input.type !== 'formula') return scalarValue(c.input);
    if (stack.size >= 128) return { error: '#LIMIT!' };
    stack.add(k);
    const childDepth = { depth: 0 };
    let result: CellValue;
    try {
      result = evaluateFormula(
        parseFormula(c.input.expression),
        (ref: Reference) => {
          const target = ref.sheet ? this.findSheet(ref.sheet) : sheetId;
          return target
            ? this.value(target, ref.row, ref.column, stack, childDepth)
            : { error: '#REF!' };
        },
        undefined,
        this.dateSystem,
      );
    } catch (e) {
      result = { error: e instanceof Error ? e.message : '#VALUE!' };
    } finally {
      stack.delete(k);
    }
    const depth = childDepth.depth + 1;
    if (parent) parent.depth = Math.max(parent.depth, depth);
    if (!(result && typeof result === 'object' && ['#LIMIT!', '#CYCLE!'].includes(result.error))) {
      // Account for keys and per-entry overhead as well as string/error payloads.
      const bytes =
        160 +
        k.length * 2 +
        (typeof result === 'string' ? result.length * 2 : 0) +
        (result && typeof result === 'object' ? result.error.length * 2 : 0);
      this.cache.set(k, { value: result, depth }, bytes);
    }
    return result;
  }
}
