import { normalizeFormulaPrecedence } from '@opensheetjs/core';

/** OOXML requires the future-function prefix for these supported modern functions. */
export function toXlsxFormula(expression: string): string {
  return expression.replace(
    /"(?:[^"]|"")*"|'(?:[^']|'')*'|(?<![\w.])(?:_xlfn\.)?(CONCAT|TEXTJOIN|IFNA|IFS|DAYS|XLOOKUP)(?=\s*\()/gi,
    (token, name: string | undefined) => (name ? `_xlfn.${name.toUpperCase()}` : token),
  );
}

/** Conservatively flag cases where OpenSheet's current power grammar can differ from Excel. */
export function mayDifferInPowerPrecedence(expression: string): boolean {
  if (!expression.includes('^')) return false;
  const powers = [0];
  let expectOperand = true;
  let prefixSign = false;
  let power = false;
  const number = /(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/y;
  for (let i = 0; i < expression.length; i++) {
    const character = expression[i];
    number.lastIndex = i;
    const numeric = expectOperand ? number.exec(expression) : null;
    if (numeric) {
      i = number.lastIndex - 1;
      expectOperand = false;
    } else if (character === '"' || character === "'") {
      const quote = character;
      while (++i < expression.length) {
        if (expression[i] !== quote) continue;
        if (expression[i + 1] === quote) i++;
        else break;
      }
      expectOperand = false;
    } else if (/\s/.test(character)) continue;
    else if (character === '(') {
      powers.push(0);
      expectOperand = true;
    } else if (character === ')') {
      if (powers.length > 1) powers.pop();
      expectOperand = false;
    } else if (character === '^') {
      power = true;
      if (++powers[powers.length - 1] > 1) return true;
      expectOperand = true;
    } else if ('+-*/&=<>:,'.includes(character)) {
      if (expectOperand && (character === '-' || character === '+')) {
        prefixSign = true;
      } else powers[powers.length - 1] = 0;
      expectOperand = true;
    } else if (character !== '%') expectOperand = false;
  }
  return power && prefixSign;
}
/** Per-conversion bounded memoization avoids retaining workbook formulas after conversion. */
export function createPrecedenceNormalizer() {
  const cache = new Map<string, string | null>();
  return (expression: string, source: 'native' | 'spreadsheet'): string | undefined => {
    const cacheable = expression.length <= 256;
    const key = cacheable ? `${source}:${expression}` : '';
    if (cacheable && cache.has(key)) return cache.get(key) ?? undefined;
    const normalized = normalizeFormulaPrecedence(expression, source);
    if (cacheable && cache.size < 256 && (normalized === undefined || normalized.length <= 1024))
      cache.set(key, normalized ?? null);
    return normalized;
  };
}
