import {
  evaluateFormula,
  parseFormula,
  transformReferences,
  functionDefinition,
  Aggregate,
  TextJoin,
  conditional,
  lookup,
  conditionalPlan,
  FormulaError,
  type AST,
  type Reference,
} from '@opensheetjs/formula';
import { assert, LIMITS, type CellValue, scalarValue, type CellRecord } from './types.js';
import { ReferenceIndex } from './reference-index.js';
import type { Selection, Rect } from './types.js';
import { ByteCache } from './storage.js';
type EvaluationWork = { remaining: number; signal?: AbortSignal };
type ScalarMemo = { generation: number; cache: ByteCache<CellValue> };
export class AsyncCalculation {
  private cache: ByteCache<{ value: CellValue; depth: number }>;
  private dependencies = new ReferenceIndex();
  private cacheGeneration = 0;
  private scalarGeneration = 0;
  private removed: string[] = [];
  private removing = false;
  get cacheBytes() {
    return this.cache.bytes;
  }
  get cacheBudget() {
    return this.cache.budget;
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
    private dateSystem: () => '1900' | '1904' = () => '1900',
  ) {
    this.cache = new ByteCache(budget, (id) => this.removeCached(id));
  }
  private removeCached(id: string) {
    this.cacheGeneration++;
    this.removed.push(id);
    if (this.removing) return;
    this.removing = true;
    try {
      while (this.removed.length) {
        const removed = this.removed.pop()!;
        this.dependencies.delete(removed);
        const [sheetId, row, column] = removed.split(':');
        for (const dependent of this.dependencies.query(sheetId, {
          startRow: +row,
          endRow: +row + 1,
          startColumn: +column,
          endColumn: +column + 1,
        }))
          this.cache.delete(dependent);
      }
    } finally {
      this.removing = false;
    }
  }
  clear() {
    this.cacheGeneration++;
    this.scalarGeneration++;
    this.cache.clear();
    this.dependencies.clear();
  }
  invalidate(ranges: Selection[]) {
    this.cacheGeneration++;
    this.scalarGeneration++;
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
    work?: EvaluationWork,
    memo?: ScalarMemo,
    parent?: { depth: number },
    checkpoint?: () => void,
  ): Promise<CellValue> {
    work?.signal?.throwIfAborted();
    if (work && --work.remaining < 0) return { error: '#LIMIT!' };
    const id = `${sheet}:${row}:${column}`;
    if (stack.has(id)) return { error: '#CYCLE!' };
    const cached = this.cache.get(id);
    if (cached !== undefined) {
      if (stack.size + cached.depth > 128) return { error: '#LIMIT!' };
      if (parent) parent.depth = Math.max(parent.depth, cached.depth);
      return cached.value;
    }
    const generation = this.cacheGeneration;
    const scalarGeneration = this.scalarGeneration;
    if (memo) {
      if (memo.generation !== scalarGeneration) {
        memo.cache.clear();
        memo.generation = scalarGeneration;
      }
      const scalar = memo.cache.get(id);
      if (scalar !== undefined) return scalar;
    }
    const cell = await this.read(sheet, row, column);
    work?.signal?.throwIfAborted();
    if (!cell || cell.input.type !== 'formula') {
      const result = cell && cell.input.type !== 'formula' ? scalarValue(cell.input) : null;
      if (
        memo &&
        scalarGeneration === this.scalarGeneration &&
        memo.generation === scalarGeneration
      ) {
        const bytes = 192 + id.length * 2 + (typeof result === 'string' ? result.length * 2 : 128);
        memo.cache.set(id, result, bytes);
      }
      return result;
    }
    if (stack.size >= 128) return { error: '#LIMIT!' };
    stack.add(id);
    const childDepth = { depth: 0 };
    let result: CellValue;
    let ast: AST | undefined;
    try {
      const persisted = await this.persistedAST?.(sheet, row, column);
      ast = persisted ?? parseFormula(cell.input.expression);
      result = await this.evaluate(sheet, ast, stack, work, memo, childDepth, checkpoint);
    } catch (error) {
      if (work?.signal?.aborted || (error as { code?: string }).code === 'ABORTED') throw error;
      result = { error: error instanceof Error ? error.message : '#VALUE!' };
    } finally {
      stack.delete(id);
    }
    const depth = childDepth.depth + 1;
    if (parent) parent.depth = Math.max(parent.depth, depth);
    return this.remember(id, sheet, row, column, ast, result, generation, depth);
  }
  /** Reuse bounded scalar references only for the lifetime of one caller-owned read. */
  readSession(minimumScratchBytes = 0, checkpoint?: () => void) {
    assert(
      Number.isSafeInteger(minimumScratchBytes) &&
        minimumScratchBytes >= 0 &&
        minimumScratchBytes <= 64 * 1024,
      'INVALID_ARGUMENT',
      'Invalid calculation scratch budget',
    );
    const memo: ScalarMemo = {
      generation: this.scalarGeneration,
      cache: new ByteCache(Math.max(minimumScratchBytes, Math.min(this.cache.budget, 64 * 1024))),
    };
    return (sheet: string, row: number, column: number) =>
      this.value(sheet, row, column, new Set(), undefined, memo, undefined, checkpoint);
  }
  /** A shared calculation budget for one immutable staged workbook view. */
  validationSession(signal?: AbortSignal, steps = 1_000_000) {
    assert(
      Number.isSafeInteger(steps) && steps > 0 && steps <= 1_000_000,
      'INVALID_ARGUMENT',
      'Invalid validation calculation budget',
    );
    const work: EvaluationWork = { remaining: steps, signal };
    // The parent cache and persisted ASTs can belong to the pre-staging revision.
    // Keep both the values and parsed formulas local to this immutable staged view.
    const staged = new AsyncCalculation(
      this.read,
      this.find,
      this.cache.budget,
      undefined,
      this.dateSystem,
    );
    const evaluateFormula = async (
      sheet: string,
      request: { expression: string; rowOffset: number; columnOffset: number },
    ): Promise<CellValue> => {
      signal?.throwIfAborted();
      if (--work.remaining < 0) return { error: '#LIMIT!' };
      try {
        const ast = transformReferences(parseFormula(request.expression), (ref) => {
          const row = ref.row + (ref.absoluteRow ? 0 : request.rowOffset);
          const column = ref.column + (ref.absoluteColumn ? 0 : request.columnOffset);
          return row < 0 || row >= LIMITS.rows || column < 0 || column >= LIMITS.columns
            ? null
            : { ...ref, row, column };
        });
        return await staged.evaluate(sheet, ast, new Set(), work);
      } catch (error) {
        if (signal?.aborted || (error as { code?: string }).code === 'ABORTED') throw error;
        return { error: error instanceof Error ? error.message : '#VALUE!' };
      }
    };
    return {
      forSheet: (sheet: string) => ({
        read: (row: number, column: number) => staged.value(sheet, row, column, new Set(), work),
        evaluateFormula: (request: Parameters<typeof evaluateFormula>[1]) =>
          evaluateFormula(sheet, request),
      }),
    };
  }
  private async evaluate(
    sheet: string,
    ast: AST,
    stack: Set<string>,
    shared?: EvaluationWork,
    memo?: ScalarMemo,
    depth?: { depth: number },
    checkpoint?: () => void,
  ) {
    const work = shared ?? { remaining: 10_000_000 };
    let textWork: { remaining: number } | undefined;
    const ref = async (r: Reference): Promise<CellValue> => {
      if (work.remaining % 1024 === 0) {
        await new Promise<void>((resolve) => setTimeout(resolve, 0));
        checkpoint?.();
      }
      work.signal?.throwIfAborted();
      if (--work.remaining < 0) return { error: '#LIMIT!' };
      const target = r.sheet ? this.find(r.sheet) : sheet;
      return target
        ? this.value(target, r.row, r.column, stack, shared, memo, depth, checkpoint)
        : { error: '#REF!' };
    };
    const literal = (value: CellValue): AST => ({ kind: 'literal', value });
    const run = async (ast: AST): Promise<CellValue> => {
      if (--work.remaining < 0) return { error: '#LIMIT!' };
      if (ast.kind === 'literal')
        return typeof ast.value === 'number' && !Number.isFinite(ast.value)
          ? { error: '#NUM!' }
          : ast.value;
      if (ast.kind === 'ref') return ref(ast);
      if (ast.kind === 'range') return { error: '#VALUE!' };
      if (ast.kind === 'unary')
        return evaluateFormula({ ...ast, value: literal(await run(ast.value)) }, () => null);
      if (ast.kind === 'binary') {
        const left = await run(ast.left);
        if (left && typeof left === 'object') return left;
        return evaluateFormula(
          { ...ast, left: literal(left), right: literal(await run(ast.right)) },
          () => null,
        );
      }
      const definition = functionDefinition(ast.name);
      if (!definition) return { error: '#NAME?' };
      if (ast.args.length < definition.minArgs || ast.args.length > definition.maxArgs)
        return { error: '#VALUE!' };
      if (definition.evaluation === 'conditional' || definition.evaluation === 'lookup') {
        try {
          const iterator = (definition.evaluation === 'lookup' ? lookup : conditional)(
            ast.name,
            ast.args,
            (textWork ??= { remaining: 1_000_000 }),
          );
          let step = iterator.next();
          while (!step.done) {
            if (work.remaining <= 0) return { error: '#LIMIT!' };
            step = iterator.next(await run(step.value));
          }
          return step.value;
        } catch (error) {
          if (error instanceof FormulaError) return { error: error.code };
          throw error;
        }
      }
      if (ast.name === 'IF') {
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
      if (ast.name === 'IFS') {
        if (ast.args.length % 2) return { error: '#VALUE!' };
        for (let i = 0; i < ast.args.length; i += 2) {
          const test = evaluateFormula(
            {
              kind: 'call',
              name: 'IF',
              args: [literal(await run(ast.args[i])), literal(true), literal(false)],
            },
            () => null,
          );
          if (test && typeof test === 'object') return test;
          if (test) return run(ast.args[i + 1]);
        }
        return { error: '#N/A' };
      }
      if (ast.name === 'IFERROR' || ast.name === 'IFNA') {
        const v = await run(ast.args[0]);
        return v &&
          typeof v === 'object' &&
          v.error !== '#LIMIT!' &&
          v.error !== '#CYCLE!' &&
          (ast.name === 'IFERROR' || v.error === '#N/A')
          ? run(ast.args[1])
          : v;
      }
      if (definition.evaluation === 'aggregate' || definition.evaluation === 'text') {
        const aggregate =
          definition.evaluation === 'aggregate'
            ? new Aggregate(ast.name)
            : ast.name === 'TEXTJOIN'
              ? new TextJoin(await run(ast.args[0]), await run(ast.args[1]))
              : new TextJoin();
        for (const arg of ast.name === 'TEXTJOIN' ? ast.args.slice(2) : ast.args) {
          if (arg.kind !== 'range') aggregate.add(await run(arg), arg.kind === 'ref');
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
                if (work.remaining <= 0) return { error: '#LIMIT!' };
                aggregate.add(await ref({ ...arg.start, row: r, column: c }), true);
              }
        }
        return aggregate.result();
      }

      const args: AST[] = [];
      for (const arg of ast.args) {
        const value = await run(arg);
        if (value && typeof value === 'object') return value;
        args.push(literal(value));
      }
      return evaluateFormula({ ...ast, args }, () => null, undefined, this.dateSystem());
    };
    return run(ast);
  }
  private remember(
    id: string,
    sheet: string,
    row: number,
    column: number,
    ast: AST | undefined,
    result: CellValue,
    generation: number,
    depth: number,
  ) {
    // Internal failures depend on the active traversal and are not stable cached values.
    if (result && typeof result === 'object' && ['#LIMIT!', '#CYCLE!'].includes(result.error))
      return result;
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
      } else if (node.kind === 'call') {
        node.args.forEach(visit);
        if (functionDefinition(node.name)?.evaluation === 'conditional') {
          try {
            const plan = conditionalPlan(node.name, node.args);
            if (plan.values)
              add(plan.values.start, {
                ...plan.values.start,
                row: plan.values.start.row + plan.values.rows - 1,
                column: plan.values.start.column + plan.values.columns - 1,
              });
          } catch {
            /* Invalid plans have no additional projected dependencies. */
          }
        }
      }
    };
    if (ast) visit(ast);
    refs.push({
      sheet,
      range: { startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 1 },
    });
    const bytes = 80 + JSON.stringify(refs).length * 2 + JSON.stringify(result).length * 2;
    if (bytes > this.cache.budget) {
      this.cacheGeneration++;
      return result;
    }
    if (generation !== this.cacheGeneration) return result;
    this.cache.set(id, { value: result, depth }, bytes);
    if (generation !== this.cacheGeneration) this.cache.delete(id);
    else this.dependencies.add(id, refs);
    return result;
  }
}
