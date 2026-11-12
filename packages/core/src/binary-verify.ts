import { readBinary } from './binary-file.js';
import { decodeBlock } from './block-codec.js';
import { assert } from './types.js';
import { verificationYield } from './verification-yield.js';

export function decodeVerifiedBinaryPage(
  bytes: Uint8Array,
  row: number,
  column: number,
  rows: Set<number>,
  columns: Set<number>,
  styles: Set<string>,
) {
  const page = decodeBlock(bytes);
  assert(
    page.row === row && page.column === column,
    'INVALID_ARGUMENT',
    'Binary page coordinates differ',
  );
  const cells = page.cells;
  for (const cell of Object.values(cells)) {
    assert(
      rows.has(cell.rowId) && columns.has(cell.columnId),
      'INVALID_ARGUMENT',
      'Cell outside sheet axes',
    );
    assert(
      cell.styleId === undefined || styles.has(cell.styleId),
      'INVALID_ARGUMENT',
      'Unknown style',
    );
  }
  return cells;
}

/** Validate a complete native binary file without changing browser storage. */
export async function verifyBinary(
  input: Blob | Uint8Array,
  options: { signal?: AbortSignal } = {},
) {
  options.signal?.throwIfAborted();
  const checkpoint = verificationYield(options.signal);
  let workbookId = '',
    sheets = 0,
    pages = 0,
    cells = 0;
  let rows: Set<number>[] = [],
    columns: Set<number>[] = [];
  let styles = new Set<string>();
  for await (const item of readBinary(input, { signal: options.signal })) {
    options.signal?.throwIfAborted();
    if ('metadata' in item) {
      workbookId = item.snapshot.workbookId;
      sheets = item.snapshot.sheets.length;
      rows = item.snapshot.sheets.map((sheet) => new Set(sheet.rowOrder));
      columns = item.snapshot.sheets.map((sheet) => new Set(sheet.columnOrder));
      styles = new Set(Object.keys(item.snapshot.styles));
      await checkpoint();
      continue;
    }
    const records = decodeVerifiedBinaryPage(
      item.bytes,
      item.row,
      item.column,
      rows[item.sheet],
      columns[item.sheet],
      styles,
    );
    pages++;
    cells += Object.keys(records).length;
    await checkpoint();
  }
  options.signal?.throwIfAborted();
  return { workbookId, sheets, pages, cells };
}
