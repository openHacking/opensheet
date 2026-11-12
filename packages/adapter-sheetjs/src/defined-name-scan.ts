/** Bounded, conservative detection of defined-name references. */
export function definedNameIndex(
  names: readonly { Name: string; Sheet?: unknown }[],
  sheetCount: number,
) {
  const global = new Set<string>();
  const local = new Map<number, Set<string>>();
  let incomplete = false;
  const limit = 8192;
  let retained = 0;
  for (const item of names) {
    if (item.Name.length > 255 || !/^[\p{L}_\\][\p{L}\p{N}\p{M}_.]*$/u.test(item.Name)) {
      incomplete = true;
      continue;
    }
    const scope = item.Sheet;
    if (
      scope !== undefined &&
      (!Number.isInteger(scope) || (scope as number) < 0 || (scope as number) >= sheetCount)
    ) {
      incomplete = true;
      continue;
    }
    const key = item.Name.toUpperCase();
    const target = scope === undefined ? global : (local.get(scope as number) ?? new Set<string>());
    if (!target.has(key)) {
      if (retained >= limit) {
        incomplete = true;
        continue;
      }
      target.add(key);
      retained++;
      if (scope !== undefined) local.set(scope as number, target);
    }
  }
  return { global, local, incomplete };
}

export function referencesDefinedName(
  expression: string,
  global: Set<string>,
  local?: Set<string>,
  qualifiedLocal?: ReadonlyMap<string, Set<string>>,
): boolean {
  if (!global.size && !local?.size && !qualifiedLocal?.size) return false;
  const identifier = /[\p{L}_\\][\p{L}\p{N}\p{M}_.]*/uy;
  const errorToken = /#(?:REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA)/iy;
  const qualifiedIdentifier = /[\p{L}_\\][\p{L}\p{N}\p{M}_.]*/uy;
  const qualified = (sheet: string, bang: number) => {
    const scope = qualifiedLocal?.get(sheet.toUpperCase());
    if (!scope) return false;
    let start = bang + 1;
    while (start < expression.length && /\s/.test(expression[start])) start++;
    qualifiedIdentifier.lastIndex = start;
    const match = qualifiedIdentifier.exec(expression);
    if (!match) return false;
    let end = qualifiedIdentifier.lastIndex;
    while (end < expression.length && /\s/.test(expression[end])) end++;
    const token = match[0];
    return (
      expression[end] !== '!' &&
      expression[end] !== '(' &&
      expression[end] !== '[' &&
      expression[end] !== '$' &&
      !/[A-Za-z_0-9.\\\u0080-\uffff]/.test(expression[end] ?? '') &&
      !/^(?:TRUE|FALSE)$/i.test(token) &&
      !/^[A-Za-z]+[1-9]\d*$/.test(token) &&
      (scope.has(token.toUpperCase()) || global.has(token.toUpperCase()))
    );
  };
  let previous = '';
  for (let i = 0; i < expression.length; ) {
    const char = expression[i];
    if (char === '#') {
      errorToken.lastIndex = i;
      const error = errorToken.exec(expression);
      if (error) {
        i += error[0].length;
        previous = '#';
        continue;
      }
    }
    if (char === '"' || char === "'") {
      const start = i + 1;
      let end = start;
      i++;
      while (i < expression.length) {
        end = expression.indexOf(char, i);
        if (end < 0) return false;
        i = end + 1;
        if (expression[i] !== char) break;
        i++;
      }
      if (char === "'" && qualifiedLocal?.size && previous !== ']') {
        let bang = i;
        while (bang < expression.length && /\s/.test(expression[bang])) bang++;
        const sheet = expression.slice(start, end).replace(/''/g, "'");
        if (!/[\[\]]/.test(sheet) && expression[bang] === '!' && qualified(sheet, bang))
          return true;
      }
      previous = char;
      continue;
    }
    if (char === '[') {
      let depth = 1;
      i++;
      while (i < expression.length && depth) {
        if (expression[i] === '[') depth++;
        else if (expression[i] === ']') depth--;
        i++;
      }
      previous = ']';
      continue;
    }
    identifier.lastIndex = i;
    const match = identifier.exec(expression);
    if (!match) {
      if (!/\s/.test(char)) previous = char;
      i++;
      continue;
    }
    i = identifier.lastIndex;
    let next = i;
    while (next < expression.length && /\s/.test(expression[next])) next++;
    const token = match[0];
    if (
      expression[next] === '!' &&
      previous !== ']' &&
      !/[A-Za-z_0-9.\\\u0080-\uffff]/.test(previous) &&
      qualified(token, next)
    )
      return true;
    if (
      previous !== '!' &&
      previous !== '$' &&
      !/[A-Za-z_0-9.\\\u0080-\uffff]/.test(previous) &&
      expression[next] !== '!' &&
      expression[next] !== '(' &&
      expression[next] !== '[' &&
      expression[next] !== '$' &&
      !/[A-Za-z_0-9.\\\u0080-\uffff]/.test(expression[next] ?? '') &&
      !/^(?:TRUE|FALSE)$/i.test(token) &&
      !/^[A-Za-z]+[1-9]\d*$/.test(token)
    ) {
      const key = token.toUpperCase();
      if (global.has(key) || local?.has(key)) return true;
    }
    previous = token.at(-1)!;
  }
  return false;
}
