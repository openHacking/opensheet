import { truth } from './coercion.js';
import { FormulaError } from './errors.js';
import { functionDefinition } from './functions.js';
import type { FormulaValue } from './types.js';

function directNumber(value: string | boolean): number | undefined {
  if (typeof value === 'boolean') return Number(value);
  const text = value.trim();
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(text)) return undefined;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : undefined;
}

/** Constant-space accumulator shared by pure and Worker-backed calculation. */
export class Aggregate {
  private sum = 0;
  private count = 0;
  private present = 0;
  private min = Infinity;
  private max = -Infinity;
  private logicalCount = 0;
  private and = true;
  private or = false;
  private error?: { error: string };
  private controlError?: { error: string };
  constructor(private readonly name: string) {
    if (functionDefinition(name)?.evaluation !== 'aggregate') throw new FormulaError('#NAME?');
  }
  add(value: FormulaValue, referenced: boolean) {
    if (value !== null) this.present++;
    if (value && typeof value === 'object') this.error ??= value;
    if (value && typeof value === 'object' && ['#LIMIT!', '#CYCLE!'].includes(value.error))
      this.controlError ??= value;
    if (typeof value === 'number') {
      this.sum += value;
      this.count++;
      this.min = Math.min(this.min, value);
      this.max = Math.max(this.max, value);
    } else if (!referenced && (typeof value === 'string' || typeof value === 'boolean'))
      this.addDirect(value);
    if (this.name === 'AND' || this.name === 'OR') {
      if (referenced && (value === null || typeof value === 'string')) return;
      try {
        const logical = truth(value);
        this.logicalCount++;
        this.and &&= logical;
        this.or ||= logical;
      } catch (error) {
        if (!(error instanceof FormulaError)) throw error;
        this.error ??= { error: error.code };
      }
    }
  }
  private addDirect(value: string | boolean) {
    if (this.name === 'SUM' || this.name === 'MIN' || this.name === 'MAX') {
      const numeric = directNumber(value);
      if (numeric === undefined) this.error ??= { error: '#VALUE!' };
      else {
        this.sum += numeric;
        this.count++;
        this.min = Math.min(this.min, numeric);
        this.max = Math.max(this.max, numeric);
      }
    } else if (
      this.name === 'AVERAGE' &&
      typeof value === 'string' &&
      directNumber(value) === undefined
    )
      this.error ??= { error: '#VALUE!' };
    else if (this.name === 'COUNT' && directNumber(value) !== undefined) this.count++;
  }
  result(): FormulaValue {
    if (this.controlError) return this.controlError;
    if (this.name === 'COUNT') return this.count;
    if (this.name === 'COUNTA') return this.present;
    if (this.error) return this.error;
    if (this.name === 'AND' || this.name === 'OR')
      return this.logicalCount ? (this.name === 'AND' ? this.and : this.or) : { error: '#VALUE!' };
    if (this.name === 'MIN') return this.count ? this.min : 0;
    if (this.name === 'MAX') return this.count ? this.max : 0;
    if (this.name === 'AVERAGE' && !this.count) return { error: '#DIV/0!' };
    const result = this.name === 'AVERAGE' ? this.sum / this.count : this.sum;
    return Number.isFinite(result) ? result : { error: '#NUM!' };
  }
}
