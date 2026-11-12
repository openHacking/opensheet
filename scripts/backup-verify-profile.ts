import { readBinary } from '../packages/core/src/binary-file.js';
import { decodeBlock } from '../packages/core/src/block-codec.js';
import { verificationYield } from '../packages/core/src/verification-yield.js';

/** Diagnostic replica, separate from the production verifier; times include async waits. */
export async function profileBinaryVerification(input: Blob) {
  const iterator = readBinary(input);
  const checkpoint = verificationYield();
  const reads: number[] = [],
    checks: number[] = [],
    decodes: number[] = [],
    expansions: number[] = [],
    references: number[] = [];
  let metadataMs = 0,
    pages = 0,
    cells = 0;
  let rows: Set<number>[] = [],
    columns: Set<number>[] = [],
    styles = new Set<string>();
  while (true) {
    const started = performance.now();
    const next = await iterator.next();
    const readMs = performance.now() - started;
    if (next.done) break;
    const item = next.value;
    if ('metadata' in item) {
      const start = performance.now();
      rows = item.snapshot.sheets.map((sheet) => new Set(sheet.rowOrder));
      columns = item.snapshot.sheets.map((sheet) => new Set(sheet.columnOrder));
      styles = new Set(Object.keys(item.snapshot.styles));
      metadataMs = readMs + performance.now() - start;
    } else {
      reads.push(readMs);
      const start = performance.now();
      const page = decodeBlock(item.bytes);
      const decodedAt = performance.now();
      if (page.row !== item.row || page.column !== item.column)
        throw new Error('Coordinates differ');
      const decoded = Object.values(page.cells);
      const expandedAt = performance.now();
      for (const cell of decoded)
        if (
          !rows[item.sheet].has(cell.rowId) ||
          !columns[item.sheet].has(cell.columnId) ||
          (cell.styleId !== undefined && !styles.has(cell.styleId))
        )
          throw new Error('Invalid cell reference');
      const checkedAt = performance.now();
      decodes.push(decodedAt - start);
      expansions.push(expandedAt - decodedAt);
      references.push(checkedAt - expandedAt);
      checks.push(checkedAt - start);
      pages++;
      cells += decoded.length;
    }
    await checkpoint();
  }
  const stats = (samples: number[]) => {
    const ordered = [...samples].sort((a, b) => a - b);
    return {
      count: ordered.length,
      totalMs: samples.reduce((sum, value) => sum + value, 0),
      p95Ms: ordered[Math.floor((ordered.length - 1) * 0.95)] ?? 0,
      maxMs: ordered.at(-1) ?? 0,
    };
  };
  return {
    pages,
    cells,
    metadataMs,
    read: stats(reads),
    validate: stats(checks),
    decode: stats(decodes),
    expand: stats(expansions),
    references: stats(references),
  };
}
