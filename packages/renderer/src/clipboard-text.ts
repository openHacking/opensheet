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
