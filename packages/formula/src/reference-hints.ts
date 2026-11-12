import { columnIndex } from './address.js';
import { FORMULA_COLUMNS, FORMULA_ROWS } from './reference-bounds.js';

export interface FormulaReferenceHint {
  sheet?: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}
export type FormulaReferenceSpan = FormulaReferenceHint & { start: number; end: number };

function* referenceSpans(source: string): Generator<FormulaReferenceSpan> {
  if (!source.startsWith('=') || source.length > 32767) return;
  const word = /[A-Za-z0-9_.$]/;
  let i = 1;
  const skipSpace = (position: number) => {
    while (position < source.length && /\s/.test(source[position]!)) position++;
    return position;
  };
  const token = (position: number) => {
    let end = position;
    while (end < source.length && word.test(source[end]!)) end++;
    return { text: source.slice(position, end), end };
  };
  const cell = (text: string) => {
    const match = /^\$?([A-Za-z]{1,3})\$?([1-9]\d{0,6})$/.exec(text);
    if (!match) return undefined;
    const row = Number(match[2]) - 1,
      column = columnIndex(match[1]!);
    return row < FORMULA_ROWS && column < FORMULA_COLUMNS ? { row, column } : undefined;
  };
  while (i < source.length) {
    const start = i;
    const ch = source[i]!;
    if (ch === '"') {
      i++;
      while (i < source.length) {
        if (source[i++] === '"') {
          if (source[i] === '"') i++;
          else break;
        }
      }
      continue;
    }
    let sheet: string | undefined;
    if (ch === "'") {
      i++;
      let name = '',
        closed = false;
      while (i < source.length) {
        const character = source[i++]!;
        if (character === "'") {
          if (source[i] === "'") {
            name += "'";
            i++;
          } else {
            closed = true;
            break;
          }
        } else name += character;
      }
      const bang = skipSpace(i);
      if (!closed || source[bang] !== '!') continue;
      sheet = name;
      i = skipSpace(bang + 1);
    } else if (!word.test(ch)) {
      // Error literals such as #REF! are not sheet names.
      if (ch === '#') {
        i++;
        while (i < source.length && /[A-Za-z0-9/!?]/.test(source[i]!)) i++;
      } else i++;
      continue;
    }
    let first = token(i);
    if (!first.text) {
      i++;
      continue;
    }
    let after = skipSpace(first.end);
    if (sheet === undefined && source[after] === '!') {
      sheet = first.text;
      i = skipSpace(after + 1);
      first = token(i);
      after = skipSpace(first.end);
    }
    const a = cell(first.text);
    i = first.end;
    if (!a || source[after] === '(') continue;
    let b = a;
    if (source[after] === ':') {
      const last = token(skipSpace(after + 1));
      const parsed = cell(last.text);
      if (parsed) {
        b = parsed;
        i = last.end;
      }
    }
    yield {
      start,
      end: i,
      ...(sheet === undefined ? {} : { sheet }),
      startRow: Math.min(a.row, b.row),
      startColumn: Math.min(a.column, b.column),
      endRow: Math.max(a.row, b.row) + 1,
      endColumn: Math.max(a.column, b.column) + 1,
    };
  }
}

/** Extract at most 64 A1 references from an incomplete draft without evaluating it. */
export function formulaReferenceSpans(source: string): FormulaReferenceSpan[] {
  const result: FormulaReferenceSpan[] = [];
  for (const span of referenceSpans(source)) {
    result.push(span);
    if (result.length === 64) break;
  }
  return result;
}

/** Find the reference containing a draft selection without retaining preceding references. */
export function formulaReferenceSpanAt(source: string, start: number, end: number) {
  if (
    !Number.isInteger(start) ||
    !Number.isInteger(end) ||
    start < 0 ||
    end < start ||
    end > source.length
  )
    return undefined;
  for (const span of referenceSpans(source)) {
    if (span.start > start) break;
    if (start >= span.start && end <= span.end) return span;
  }
  return undefined;
}

export function formulaReferenceHints(source: string): FormulaReferenceHint[] {
  return formulaReferenceSpans(source).map(({ start: _start, end: _end, ...hint }) => hint);
}
