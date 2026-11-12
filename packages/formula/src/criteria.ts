import { checked } from './coercion.js';
import { fail } from './errors.js';
import type { FormulaValue } from './types.js';

const numeric = (value: string) =>
  /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(value.trim()) &&
  Number.isFinite(Number(value));
type Token = { kind: 'star' | 'one' | 'literal'; value?: string };
export function wildcard(pattern: string, work = { remaining: 1_000_000 }) {
  if (pattern.length > 32767) fail('#LIMIT!');
  work.remaining -= pattern.length;
  if (work.remaining < 0) fail('#LIMIT!');
  const tokens: Token[] = [];
  const points = Array.from(pattern.toLowerCase());
  for (let i = 0; i < points.length; i++) {
    const c = points[i];
    if (c === '~' && i + 1 < points.length && '*?~'.includes(points[i + 1]))
      tokens.push({ kind: 'literal', value: points[++i] });
    else if (c === '*' && tokens.at(-1)?.kind === 'star') continue;
    else
      tokens.push(
        c === '*' ? { kind: 'star' } : c === '?' ? { kind: 'one' } : { kind: 'literal', value: c },
      );
  }
  return (text: string) => {
    work.remaining -= text.length;
    if (work.remaining < 0) fail('#LIMIT!');
    text = text.toLowerCase();
    let at = 0,
      token = 0,
      star = -1,
      retry = 0;
    const next = (index: number) => index + (text.codePointAt(index)! > 0xffff ? 2 : 1);
    while (at < text.length) {
      if (--work.remaining < 0) fail('#LIMIT!');
      const t = tokens[token];
      if (t?.kind === 'star') {
        star = token++;
        retry = at;
      } else if (t && (t.kind === 'one' || t.value === text.slice(at, next(at)))) {
        at = next(at);
        token++;
      } else if (star >= 0 && retry < text.length) {
        retry = next(retry);
        at = retry;
        token = star + 1;
      } else return false;
    }
    while (tokens[token]?.kind === 'star') token++;
    return token === tokens.length;
  };
}

/** Compile once per criterion, never build an untrusted backtracking regular expression. */
export function criterion(
  input: FormulaValue,
  work = { remaining: 1_000_000 },
): (value: FormulaValue) => boolean {
  checked(input);
  let op = '=',
    target: string | number | boolean = input === null ? 0 : (input as string | number | boolean);
  if (typeof target === 'string') {
    if (target.length > 32767) fail('#LIMIT!');
    work.remaining -= target.length;
    if (work.remaining < 0) fail('#LIMIT!');
    const match = /^(<=|>=|<>|=|<|>)(.*)$/s.exec(target);
    if (match) {
      op = match[1];
      target = match[2];
    }
    if (numeric(target)) target = Number(target);
  }
  const textTarget = typeof target === 'string' ? target.toLowerCase() : undefined;
  const pattern =
    textTarget !== undefined && /[?*~]/.test(textTarget) && (op === '=' || op === '<>')
      ? wildcard(textTarget, work)
      : undefined;
  return (value) => {
    if (value && typeof value === 'object') {
      if (value.error === '#LIMIT!' || value.error === '#CYCLE!') return fail(value.error);
      if (textTarget?.startsWith('#') && (op === '=' || op === '<>'))
        return op === '='
          ? value.error.toLowerCase() === textTarget
          : value.error.toLowerCase() !== textTarget;
      return false;
    }
    // String normalization and numeric parsing touch the whole input even when
    // a comparison or wildcard match can reject after its first character.
    if (typeof value === 'string' && typeof target !== 'boolean' && !pattern) {
      work.remaining -= value.length;
      if (work.remaining < 0) fail('#LIMIT!');
    }
    let comparison: number;
    if (typeof target === 'number') {
      const v =
        value === null ? 0 : typeof value === 'string' && numeric(value) ? Number(value) : value;
      if (typeof v !== 'number') return op === '<>';
      comparison = v === target ? 0 : v < target ? -1 : 1;
    } else if (typeof target === 'boolean') {
      if (typeof value !== 'boolean') return op === '<>';
      comparison = Number(value) - Number(target);
    } else {
      if (pattern) {
        const matches = typeof value === 'string' && pattern(value);
        return op === '=' ? matches : !matches;
      }
      const v = value === null ? '' : value;
      if (typeof v !== 'string') return op === '<>';
      const lower = v.toLowerCase();
      comparison = lower === textTarget ? 0 : lower < textTarget! ? -1 : 1;
    }
    return op === '='
      ? comparison === 0
      : op === '<>'
        ? comparison !== 0
        : op === '<'
          ? comparison < 0
          : op === '>'
            ? comparison > 0
            : op === '<='
              ? comparison <= 0
              : comparison >= 0;
  };
}
