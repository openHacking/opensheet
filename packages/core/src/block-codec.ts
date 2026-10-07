import { assert, key, type CellInput, type CellRecord } from './types.js';
import { BLOCK_COLUMNS } from './storage.js';
export type EncodedBlock = {
  rows: string[];
  columns: string[];
  cells: Array<[number, number, unknown, Omit<CellRecord, 'rowId' | 'columnId' | 'input'>?]>;
  dependencies: Record<string, unknown>;
};
const tags = { blank: 0, number: 1, string: 2, boolean: 3, error: 4, formula: 5 } as const;
export function encodeBlock(
  cells: Record<string, CellRecord>,
  rowIds: string[],
  columnIds: string[],
  dependencies: Record<string, unknown>,
): EncodedBlock {
  const rows = new Map(rowIds.map((id, i) => [id, i])),
    columns = new Map(columnIds.map((id, i) => [id, i]));
  const encoded: EncodedBlock['cells'] = [];
  for (const cell of Object.values(cells)) {
    const { rowId, columnId, input, ...extra } = cell;
    const index = rows.get(rowId)! * BLOCK_COLUMNS + columns.get(columnId)!;
    const value =
      'value' in input
        ? input.value
        : input.type === 'error'
          ? input.code
          : input.type === 'formula'
            ? input.expression
            : null;
    const entry: EncodedBlock['cells'][number] = [index, tags[input.type], value];
    if (Object.keys(extra).length) entry.push(extra);
    encoded.push(entry);
  }
  return { rows: rowIds, columns: columnIds, cells: encoded, dependencies };
}
export function decodeBlock(encoded: EncodedBlock) {
  const cells: Record<string, CellRecord> = {};
  for (const [offset, tag, value, extra] of encoded.cells) {
    const rowId = encoded.rows[Math.floor(offset / BLOCK_COLUMNS)],
      columnId = encoded.columns[offset % BLOCK_COLUMNS];
    assert(
      rowId && columnId && Number.isInteger(offset) && offset >= 0,
      'CORRUPT_STORAGE',
      'Invalid block offset',
    );
    let input: CellInput;
    switch (tag) {
      case 0:
        input = { type: 'blank' };
        break;
      case 1:
        input = { type: 'number', value: value as number };
        break;
      case 2:
        input = { type: 'string', value: value as string };
        break;
      case 3:
        input = { type: 'boolean', value: value as boolean };
        break;
      case 4:
        input = { type: 'error', code: value as string };
        break;
      case 5:
        input = { type: 'formula', expression: value as string };
        break;
      default:
        throw new Error('Invalid stored cell type');
    }
    cells[key(rowId, columnId)] = { rowId, columnId, input, ...extra };
  }
  return { cells, dependencies: encoded.dependencies };
}
