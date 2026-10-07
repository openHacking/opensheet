export function stringifyDelimited(
  rows: unknown[][],
  separator = ',',
  safe = true,
): { text: string; sanitized: number } {
  let sanitized = 0;
  const text = rows
    .map((row) =>
      row
        .map((value) => {
          let s = String(value ?? '');
          if (safe && typeof value === 'string' && /^[\s]*[=+@-]|^[\t\r]/.test(s)) {
            s = "'" + s;
            sanitized++;
          }
          return s.includes(separator) || /["\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        })
        .join(separator),
    )
    .join('\r\n');
  return { text, sanitized };
}
export function parseDelimited(text: string, separator = ','): string[][] {
  if (text.length > 20 * 1024 * 1024) throw new Error('Text file exceeds 20 MiB');
  const rows: string[][] = [];
  let row: string[] = [],
    cell = '',
    quoted = false,
    afterQuote = false;
  const input = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < input.length; i++) {
    const c = input[i];
    if (quoted) {
      if (c === '"') {
        if (input[i + 1] === '"') {
          cell += '"';
          i++;
        } else {
          quoted = false;
          afterQuote = true;
        }
      } else cell += c;
    } else if (c === '"' && !cell && !afterQuote) quoted = true;
    else if (c === separator) {
      row.push(cell);
      cell = '';
      afterQuote = false;
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && input[i + 1] === '\n') i++;
      row.push(cell);
      rows.push(row);
      cell = '';
      row = [];
      afterQuote = false;
    } else {
      if (afterQuote) throw new Error('Unexpected character after quoted field');
      cell += c;
    }
    if (cell.length > 32768 || rows.length > 100000 || row.length > 1000)
      throw new Error('Delimited data exceeds limits');
  }
  if (quoted) throw new Error('Unterminated quoted field');
  if (cell || row.length || !rows.length) {
    row.push(cell);
    rows.push(row);
  }
  const width = rows.reduce((max, r) => Math.max(max, r.length), 0);
  if (rows.length * width > 200000) throw new Error('Too many cells');
  return rows.map((r) => [...r, ...Array(width - r.length).fill('')]);
}
