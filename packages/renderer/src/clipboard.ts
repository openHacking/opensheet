import {
  OpenSheetError,
  parseInput,
  type Range,
  type Selection,
  type Workbook,
} from '@opensheetjs/core';
export async function copyText(range: Range) {
  return (await range.getDisplayValues())
    .map((row) => row.map((v) => (/[\t\n"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v)).join('\t'))
    .join('\n');
}

export async function pasteText(
  book: Workbook,
  selected: Selection,
  text: string,
): Promise<Selection> {
  if (text.length > 2 * 1024 * 1024)
    throw new OpenSheetError('LIMIT_EXCEEDED', 'Paste is too large');
  const rows: string[][] = [];
  let row: string[] = [],
    cell = '',
    quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === '"' && (quoted || !cell)) {
      if (quoted && text[i + 1] === '"') {
        cell += '"';
        i++;
      } else quoted = !quoted;
    } else if (!quoted && (c === '\t' || c === '\n' || c === '\r')) {
      row.push(cell);
      cell = '';
      if (c !== '\t') {
        rows.push(row);
        row = [];
        if (c === '\r' && text[i + 1] === '\n') i++;
      }
    } else cell += c;
  }
  if (quoted) throw new OpenSheetError('INVALID_ARGUMENT', 'Unclosed quote');
  if (cell || row.length || !rows.length) {
    row.push(cell);
    rows.push(row);
  }
  const width = Math.max(...rows.map((r) => r.length));
  const cells = rows.flatMap((row, i) =>
    Array.from({ length: width }, (_, j) => ({
      row: selected.startRow + i,
      column: selected.startColumn + j,
      input: parseInput(row[j] ?? ''),
    })),
  );
  await book.execute({ type: 'core.cells.set', payload: { sheetId: selected.sheetId, cells } });
  return {
    ...selected,
    endRow: selected.startRow + rows.length,
    endColumn: selected.startColumn + width,
  };
}
