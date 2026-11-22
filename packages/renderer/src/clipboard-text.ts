/** Encode a rectangular display matrix as tab-separated clipboard text. */
export function encodeClipboardText(rows: readonly (readonly string[])[]): string {
  return rows
    .map((row, rowIndex) =>
      row
        .map((value) =>
          /[\t\r\n"]/.test(value) ||
          (rows.length > 1 && rowIndex === rows.length - 1 && row.length === 1 && value === '')
            ? `"${value.replace(/"/g, '""')}"`
            : value,
        )
        .join('\t'),
    )
    .join('\n');
}

/** Exact encoded size for a field, excluding row and column separators. */
export function clipboardTextCellLength(value: string, quoteEmpty: boolean): number {
  if (!quoteEmpty && !/[\t\r\n"]/.test(value)) return value.length;
  let quotes = 0;
  for (
    let position = value.indexOf('"');
    position >= 0;
    position = value.indexOf('"', position + 1)
  )
    quotes++;
  return value.length + quotes + 2;
}
