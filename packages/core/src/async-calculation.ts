import { evaluateFormula, parseFormula, type AST, type Reference } from '@opensheetjs/formula';
import { type CellValue, scalarValue, type CellRecord } from './types.js';
import { ByteCache } from './storage.js';
export class AsyncCalculation {
  private cache: ByteCache<CellValue>;
  get cacheBytes() {
    return this.cache.bytes;
  }
  constructor(
    private read: (sheet: string, row: number, column: number) => Promise<CellRecord | undefined>,
    private find: (name: string) => string | undefined,
    budget = 4 * 1024 * 1024,
    private persistedAST?: (
      sheet: string,
      row: number,
      column: number,
    ) => Promise<AST | null | undefined>,
  ) {
    this.cache = new ByteCache(budget);
  }
  clear() {
    this.cache.clear();
  }
  async value(
    sheet: string,
    row: number,
    column: number,
    stack = new Set<string>(),
  ): Promise<CellValue> {
    const id = `${sheet}:${row}:${column}`;
    const cached = this.cache.get(id);
    if (cached !== undefined) return cached;
    if (stack.has(id)) return { error: '#CYCLE!' };
    if (stack.size >= 128) return { error: '#LIMIT!' };
    const cell = await this.read(sheet, row, column);
    if (!cell) return null;
    if (cell.input.type !== 'formula') return scalarValue(cell.input);
    stack.add(id);
    let budget = 10_000_000;
    const ref = async (r: Reference): Promise<CellValue> => {
      if (budget % 1024 === 0) await new Promise<void>((resolve) => setTimeout(resolve, 0));
      if (--budget < 0) return { error: '#LIMIT!' };
      const target = r.sheet ? this.find(r.sheet) : sheet;
      return target ? this.value(target, r.row, r.column, stack) : { error: '#REF!' };
    };
    const literal = (value: CellValue): AST => ({ kind: 'literal', value });
    const run = async (ast: AST): Promise<CellValue> => {
      if (--budget < 0) return { error: '#LIMIT!' };
      if (ast.kind === 'literal') return ast.value;
      if (ast.kind === 'ref') return ref(ast);
      if (ast.kind === 'range') return { error: '#VALUE!' };
      if (ast.kind === 'unary')
        return evaluateFormula({ ...ast, value: literal(await run(ast.value)) }, () => null);
      if (ast.kind === 'binary')
        return evaluateFormula(
          { ...ast, left: literal(await run(ast.left)), right: literal(await run(ast.right)) },
          () => null,
        );
      if (ast.name === 'IF') {
        if (ast.args.length < 2 || ast.args.length > 3) return { error: '#VALUE!' };
        const test = evaluateFormula(
          {
            kind: 'call',
            name: 'IF',
            args: [literal(await run(ast.args[0])), literal(true), literal(false)],
          },
          () => null,
        );
        return typeof test === 'object' && test
          ? test
          : run(ast.args[test ? 1 : 2] ?? literal(false));
      }
      if (ast.name === 'IFERROR') {
        if (ast.args.length !== 2) return { error: '#VALUE!' };
        const v = await run(ast.args[0]);
        return v && typeof v === 'object' ? run(ast.args[1]) : v;
      }
      if (['SUM', 'AVERAGE', 'MIN', 'MAX', 'COUNT', 'COUNTA', 'AND', 'OR'].includes(ast.name)) {
        if (!ast.args.length || ast.args.length > 255) return { error: '#VALUE!' };
        let sum = 0,
          count = 0,
          present = 0,
          min = Infinity,
          max = -Infinity,
          and = true,
          or = false;
        let error: CellValue | undefined;
        const consume = (v: CellValue) => {
          if (v !== null) present++;
          if (v && typeof v === 'object') error = v;
          if (typeof v === 'number') {
            sum += v;
            count++;
            min = Math.min(min, v);
            max = Math.max(max, v);
          }
          if (ast.name === 'AND' || ast.name === 'OR') {
            const truth = evaluateFormula(
              { kind: 'call', name: 'NOT', args: [literal(v)] },
              () => null,
            );
            if (truth && typeof truth === 'object') error = truth;
            else {
              and &&= !truth;
              or ||= !truth;
            }
          }
        };
        for (const arg of ast.args) {
          if (arg.kind !== 'range') consume(await run(arg));
          else
            for (
              let r = Math.min(arg.start.row, arg.end.row);
              r <= Math.max(arg.start.row, arg.end.row);
              r++
            )
              for (
                let c = Math.min(arg.start.column, arg.end.column);
                c <= Math.max(arg.start.column, arg.end.column);
                c++
              ) {
                if (budget <= 0) return { error: '#LIMIT!' };
                consume(await ref({ ...arg.start, row: r, column: c }));
              }
        }
        if (ast.name === 'COUNT') return count;
        if (ast.name === 'COUNTA') return present;
        if (error !== undefined) return error;
        if (ast.name === 'SUM') return Number.isFinite(sum) ? sum : { error: '#NUM!' };
        if (ast.name === 'AVERAGE') return count ? sum / count : { error: '#DIV/0!' };
        if (ast.name === 'MIN') return count ? min : 0;
        if (ast.name === 'MAX') return count ? max : 0;
        return ast.name === 'AND' ? and : or;
      }
      const args: AST[] = [];
      for (const arg of ast.args) args.push(literal(await run(arg)));
      return evaluateFormula({ ...ast, args }, () => null);
    };
    let result: CellValue;
    try {
      const persisted = await this.persistedAST?.(sheet, row, column);
      result = await run(persisted ?? parseFormula(cell.input.expression));
    } catch (error) {
      if ((error as { code?: string }).code === 'ABORTED') throw error;
      result = { error: error instanceof Error ? error.message : '#VALUE!' };
    } finally {
      stack.delete(id);
    }
    this.cache.set(id, result);
    return result;
  }
}
