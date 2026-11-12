import { checked, number, truth } from './coercion.js';
import { fail, FormulaError } from './errors.js';
import type { FormulaValue } from './types.js';

export const FORMULA_TEXT_LIMIT = 32767;
export function formulaText(value: FormulaValue): string {
  checked(value);
  return value === null
    ? ''
    : typeof value === 'boolean'
      ? value
        ? 'TRUE'
        : 'FALSE'
      : String(value);
}
const bounded = (value: string) => (value.length <= FORMULA_TEXT_LIMIT ? value : fail());
function length(value: string) {
  let count = 0;
  for (const _point of value) count++;
  return count;
}
function slice(value: string, start: number, count: number) {
  if (!count) return '';
  let index = 0,
    offset = 0,
    begin: number | undefined;
  for (const point of value) {
    if (index === start) begin = offset;
    if (index === start + count) return bounded(value.slice(begin, offset));
    offset += point.length;
    index++;
  }
  return begin === undefined ? '' : bounded(value.slice(begin));
}
const characters = (value: FormulaValue) => {
  const count = number(value);
  if (count < 0) fail();
  return Math.trunc(count);
};

export function evaluateText(name: string, values: FormulaValue[]): FormulaValue | undefined {
  if (!['LEFT', 'RIGHT', 'MID', 'LEN', 'TRIM', 'UPPER', 'LOWER'].includes(name)) return undefined;
  const value = formulaText(values[0]);
  if (name === 'LEN') return length(value);
  if (name === 'TRIM') return bounded(value.replace(/ +/g, ' ').replace(/^ | $/g, ''));
  if (name === 'UPPER') return bounded(value.toUpperCase());
  if (name === 'LOWER') return bounded(value.toLowerCase());
  if (name === 'MID') {
    const start = number(values[1]);
    if (start < 1) fail();
    return slice(value, Math.trunc(start) - 1, characters(values[2]));
  }
  const count = values.length > 1 ? characters(values[1]) : 1;
  return slice(value, name === 'LEFT' ? 0 : Math.max(0, length(value) - count), count);
}

/** Streaming text aggregation with bounded output, shared by both evaluator paths. */
export class TextJoin {
  private value = '';
  private hasValue = false;
  private size = 0;
  private delimiter = '';
  private ignoreEmpty = false;
  private error?: { error: string };
  constructor(delimiter: FormulaValue = '', ignoreEmpty: FormulaValue = false) {
    const internal = [delimiter, ignoreEmpty].find(
      (value) => value && typeof value === 'object' && ['#LIMIT!', '#CYCLE!'].includes(value.error),
    );
    if (internal && typeof internal === 'object') {
      this.error = internal;
      return;
    }
    try {
      this.delimiter = formulaText(delimiter);
      this.ignoreEmpty = truth(ignoreEmpty);
    } catch (error) {
      if (!(error instanceof FormulaError)) throw error;
      this.error = { error: error.code };
    }
  }
  add(value: FormulaValue, _referenced = false) {
    // Internal failures must not be hidden by an earlier ordinary error.
    if (value && typeof value === 'object' && ['#LIMIT!', '#CYCLE!'].includes(value.error)) {
      if (!this.error || !['#LIMIT!', '#CYCLE!'].includes(this.error.error)) this.error = value;
      this.value = '';
      return;
    }
    if (this.error) return;
    try {
      const text = formulaText(value);
      if (this.ignoreEmpty && !text) return;
      const size = this.size + text.length + (this.hasValue ? this.delimiter.length : 0);
      if (size > FORMULA_TEXT_LIMIT) fail();
      if (this.hasValue && this.delimiter) this.value += this.delimiter;
      if (text) this.value += text;
      this.hasValue = true;
      this.size = size;
    } catch (error) {
      if (!(error instanceof FormulaError)) throw error;
      this.error = { error: error.code };
      this.value = '';
    }
  }
  result(): FormulaValue {
    return this.error ?? this.value;
  }
}
