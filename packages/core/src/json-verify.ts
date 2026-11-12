import { decodeFileCells, decodeFileMetadata, type WorkbookFile } from './file-codec.js';
import { parseWorkbookJSON } from './json-stream.js';
import { assert } from './types.js';
import { verificationYield } from './verification-yield.js';

/** Validate a replayable native JSON backup without changing browser storage. */
export async function verifyJSON(input: string | Blob, options: { signal?: AbortSignal } = {}) {
  options.signal?.throwIfAborted();
  assert(
    typeof input === 'string' || input instanceof Blob,
    'INVALID_ARGUMENT',
    'JSON verification requires a replayable string or Blob',
  );
  const checkpoint = verificationYield(options.signal);
  let metadata: WorkbookFile | undefined;
  for await (const item of parseWorkbookJSON(input, { checkpoint, signal: options.signal })) {
    options.signal?.throwIfAborted();
    if ('metadata' in item) metadata = item.metadata;
    await checkpoint();
  }
  assert(metadata, 'INVALID_ARGUMENT', 'Missing workbook metadata');
  const snapshot = decodeFileMetadata(metadata);
  const rows = snapshot.sheets.map((sheet) => new Set(sheet.rowOrder));
  const columns = snapshot.sheets.map((sheet) => new Set(sheet.columnOrder));
  const styles = Object.keys(snapshot.styles);
  const seen = new Set<string>();
  let pages = 0,
    cells = 0;
  for await (const item of parseWorkbookJSON(input, { checkpoint, signal: options.signal })) {
    options.signal?.throwIfAborted();
    if ('metadata' in item) {
      await checkpoint();
      continue;
    }
    assert(
      Number.isInteger(item.sheet) && !!snapshot.sheets[item.sheet],
      'INVALID_ARGUMENT',
      'Unknown imported sheet',
    );
    const address = `${item.sheet}/${item.block.row}/${item.block.column}`;
    assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate imported block');
    seen.add(address);
    const validated = decodeFileCells(item.block, styles, rows[item.sheet], columns[item.sheet]);
    pages++;
    cells += Object.keys(validated).length;
    await checkpoint();
  }
  options.signal?.throwIfAborted();
  return { workbookId: snapshot.workbookId, sheets: snapshot.sheets.length, pages, cells };
}
