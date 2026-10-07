import { evaluateFormula, parseFormula, type AST, type Reference } from '@opensheetjs/formula';
import { type CellValue, scalarValue, type CellRecord } from './types.js';
import { ReferenceIndex } from './reference-index.js';
import type { Selection, Rect } from './types.js';
import { ByteCache } from './storage.js';
export class AsyncCalculation {
  private cache: ByteCache<CellValue>;
  private dependencies = new ReferenceIndex();
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
    this.cache = new ByteCache(budget, (id) => this.dependencies.delete(id));
  }
  clear() {
    this.cache.clear();
    this.dependencies.clear();
  }
  invalidate(ranges: Selection[]) {
    const queue = [...ranges],
      seen = new Set<string>();
    while (queue.length) {
      const range = queue.pop()!;
      for (const id of this.dependencies.query(range.sheetId, range)) {
        if (seen.has(id)) continue;
        seen.add(id);
        this.cache.delete(id);
        const [sheetId, row, column] = id.split(':');
        queue.push({
          sheetId,
          startRow: +row,
          endRow: +row + 1,
          startColumn: +column,
          endColumn: +column + 1,
        });
      }
    }
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
    let ast: AST | undefined;
    try {
      const persisted = await this.persistedAST?.(sheet, row, column);
      ast = persisted ?? parseFormula(cell.input.expression);
      result = await run(ast);
    } catch (error) {
      if ((error as { code?: string }).code === 'ABORTED') throw error;
      result = { error: error instanceof Error ? error.message : '#VALUE!' };
    } finally {
      stack.delete(id);
    }
    const refs: Array<{ sheet: string; range: Rect }> = [];
    const add = (ref: Reference, end = ref) => {
      const target = ref.sheet ? this.find(ref.sheet) : sheet;
      if (target)
        refs.push({
          sheet: target,
          range: {
            startRow: Math.min(ref.row, end.row),
            endRow: Math.max(ref.row, end.row) + 1,
            startColumn: Math.min(ref.column, end.column),
            endColumn: Math.max(ref.column, end.column) + 1,
          },
        });
    };
    const visit = (node: AST) => {
      if (node.kind === 'ref') add(node);
      else if (node.kind === 'range') add(node.start, node.end);
      else if (node.kind === 'unary') visit(node.value);
      else if (node.kind === 'binary') {
        visit(node.left);
        visit(node.right);
      } else if (node.kind === 'call') node.args.forEach(visit);
    };
    if (ast) visit(ast);
    refs.push({
      sheet,
      range: { startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 1 },
    });
    const bytes = 64 + JSON.stringify(refs).length * 2 + JSON.stringify(result).length * 2;
    this.cache.set(id, result, bytes);
    if (bytes <= this.cache.budget) this.dependencies.add(id, refs);
    return result;
  }
}
