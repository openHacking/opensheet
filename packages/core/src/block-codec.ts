import { parseFormula } from '@opensheetjs/formula';
import { assert, key, type CellRecord, type CellInput } from './types.js';
import { encodeFileCells, decodeFileCells, type FileBlockData } from './file-codec.js';
export type EncodedBlock = Uint8Array;
const encoder = new TextEncoder();
const decoder = new TextDecoder('utf-8', { fatal: true });
const kinds = ['numbers', 'strings', 'booleans', 'formulas', 'errors'] as const;
type Kind = (typeof kinds)[number];
type Section = {
  name: Kind;
  positions: { mode: number; count: number; offset: number; bytes: number };
  type: number;
  offset: number;
  bytes: number;
};
const align = (n: number) => Math.ceil(n / 8) * 8;
const bitmap = (indices: number[], length: number) => {
  const bytes = new Uint8Array(Math.ceil(length / 8));
  for (const i of indices) bytes[i >> 3] |= 1 << (i & 7);
  return bytes;
};
function positions(indices: number[]) {
  const runs: number[] = [];
  for (const i of indices) {
    if (runs.length && runs.at(-2)! + runs.at(-1)! === i) runs[runs.length - 1]++;
    else runs.push(i, 1);
  }
  const little = (values: number[]) => {
    const bytes = new Uint8Array(values.length * 2),
      view = new DataView(bytes.buffer);
    values.forEach((value, i) => view.setUint16(i * 2, value, true));
    return bytes;
  };
  const list = little(indices),
    run = little(runs),
    bits = bitmap(indices, 2048);
  if (run.byteLength < list.byteLength && run.byteLength <= bits.byteLength)
    return { mode: 1, bytes: run };
  if (bits.byteLength < list.byteLength) return { mode: 2, bytes: bits };
  return { mode: 0, bytes: list };
}
function numbers(values: number[]) {
  const integers = values.every((v) => Number.isInteger(v) && !Object.is(v, -0));
  let data: Int8Array | Int16Array | Int32Array | Float64Array;
  if (integers && values.every((v) => v >= -128 && v <= 127)) data = new Int8Array(values);
  else if (integers && values.every((v) => v >= -32768 && v <= 32767))
    data = new Int16Array(values);
  else if (integers && values.every((v) => v >= -2147483648 && v <= 2147483647))
    data = new Int32Array(values);
  else data = new Float64Array(values);
  // Persisted numeric bytes are little endian on every platform.
  const bytes = new Uint8Array(data.byteLength),
    view = new DataView(bytes.buffer);
  values.forEach((v, i) => {
    if (data.BYTES_PER_ELEMENT === 1) view.setInt8(i, v);
    else if (data.BYTES_PER_ELEMENT === 2) view.setInt16(i * 2, v, true);
    else if (data.BYTES_PER_ELEMENT === 4) view.setInt32(i * 4, v, true);
    else view.setFloat64(i * 8, v, true);
  });
  return { type: data.BYTES_PER_ELEMENT, bytes };
}
const utf16 = (text: string) => {
  const bytes = new Uint8Array(text.length * 2),
    view = new DataView(bytes.buffer);
  for (let i = 0; i < text.length; i++) view.setUint16(i * 2, text.charCodeAt(i), true);
  return bytes;
};
function strings(values: string[]) {
  const wide = values.some((v) =>
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/u.test(v),
  );
  const unique = [...new Set(values)],
    encode = wide ? utf16 : (s: string) => encoder.encode(s);
  const pack = (texts: string[], dictionary: boolean) => {
    const parts = texts.map(encode),
      arenaBytes = parts.reduce((n, b) => n + b.byteLength, 0);
    const header = 4 + (dictionary ? values.length * 2 : 0) + (texts.length + 1) * 4;
    const bytes = new Uint8Array(header + arenaBytes),
      view = new DataView(bytes.buffer);
    view.setUint32(0, texts.length, true);
    let cursor = 4;
    if (dictionary) {
      const map = new Map(texts.map((v, i) => [v, i]));
      values.forEach((v, i) => view.setUint16(cursor + i * 2, map.get(v)!, true));
      cursor += values.length * 2;
    }
    let offset = 0;
    parts.forEach((part, i) => {
      view.setUint32(cursor + i * 4, offset, true);
      bytes.set(part, header + offset);
      offset += part.byteLength;
    });
    view.setUint32(cursor + texts.length * 4, offset, true);
    return { type: (wide ? 16 : 0) + (dictionary ? 1 : 0), bytes };
  };
  const plain = pack(values, false),
    dict = pack(unique, true);
  return dict.bytes.byteLength < plain.bytes.byteLength ? dict : plain;
}
export function encodeBlock(
  cells: Record<string, CellRecord>,
  rowIds: number[],
  columnIds: number[],
  _dependencies: Record<string, unknown>,
): EncodedBlock {
  const values = Object.values(cells),
    row = Math.floor((values[0]?.rowId ?? rowIds[0] ?? 0) / 64),
    column = Math.floor((values[0]?.columnId ?? columnIds[0] ?? 0) / 32);
  assert(
    values.every((c) => Math.floor(c.rowId / 64) === row && Math.floor(c.columnId / 32) === column),
    'INVALID_ARGUMENT',
    'Cells cross page boundaries',
  );
  const styles = [...new Set(values.flatMap((c) => (c.styleId === undefined ? [] : [c.styleId])))];
  const data = encodeFileCells(values, new Map(styles.map((s, i) => [s, i])));
  const sections: Section[] = [],
    buffers: Uint8Array[] = [];
  let offset = 0;
  for (const name of kinds) {
    const col = data[name];
    if (!col) continue;
    const pos = positions(col.positions);
    const vals =
      name === 'numbers'
        ? numbers(
            (col.values as number[]).map((v, i) =>
              data.negativeZero?.includes(col.positions[i]) ? -0 : v,
            ),
          )
        : name === 'booleans'
          ? {
              type: 0,
              bytes: bitmap(
                col.values.flatMap((v, i) => (v ? [i] : [])),
                col.values.length,
              ),
            }
          : strings(col.values as string[]);
    const section: Section = {
      name,
      positions: { mode: pos.mode, count: col.positions.length, offset, bytes: pos.bytes.length },
      type: vals.type,
      offset: 0,
      bytes: vals.bytes.length,
    };
    buffers.push(pos.bytes);
    offset += align(pos.bytes.length);
    section.offset = offset;
    buffers.push(vals.bytes);
    offset += align(vals.bytes.length);
    sections.push(section);
    delete data[name];
  }
  delete data.negativeZero;
  const meta = encoder.encode(JSON.stringify({ row, column, styles, data, sections }));
  assert(
    align(8 + meta.length) + offset <= 32 * 1024 * 1024,
    'LIMIT_EXCEEDED',
    'Page exceeds 32 MiB working budget',
  );
  const header = align(8 + meta.length),
    result = new Uint8Array(header + offset),
    view = new DataView(result.buffer);
  result.set([79, 83, 80, 3]);
  view.setUint32(4, meta.length, true);
  result.set(meta, 8);
  let cursor = header;
  for (const bytes of buffers) {
    result.set(bytes, cursor);
    cursor += align(bytes.length);
  }
  return result;
}
export class ColumnPage {
  readonly dependencies: Record<string, unknown> = {};
  readonly row: number;
  readonly column: number;
  readonly styles: string[];
  private extra: Record<string, CellRecord>;
  private columns: Array<{
    name: Kind;
    positions: Uint16Array;
    value: (index: number) => unknown;
  }> = [];
  readonly bytes: number;
  constructor(readonly buffer: Uint8Array) {
    assert(
      buffer.length >= 8 &&
        buffer[0] === 79 &&
        buffer[1] === 83 &&
        buffer[2] === 80 &&
        buffer[3] === 3,
      'CORRUPT_STORAGE',
      'Unsupported page format',
    );
    const view = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength),
      size = view.getUint32(4, true),
      base = align(8 + size);
    assert(base <= buffer.length, 'CORRUPT_STORAGE', 'Invalid page header');
    const meta = JSON.parse(decoder.decode(buffer.subarray(8, 8 + size)));
    this.row = meta.row;
    this.column = meta.column;
    this.styles = meta.styles;
    const data: FileBlockData = meta.data;
    let previousEnd = 0;
    for (const section of meta.sections as Section[]) {
      assert(
        kinds.includes(section.name) &&
          !this.columns.some((c) => c.name === section.name) &&
          section.positions.count > 0 &&
          section.positions.count <= 2048,
        'CORRUPT_STORAGE',
        'Invalid page section',
      );
      const bounds = (offset: number, bytes: number) => {
        assert(
          Number.isInteger(offset) &&
            Number.isInteger(bytes) &&
            offset >= previousEnd &&
            bytes >= 0 &&
            base + offset + bytes <= buffer.length,
          'CORRUPT_STORAGE',
          'Invalid page offsets',
        );
        previousEnd = offset + bytes;
      };
      bounds(section.positions.offset, section.positions.bytes);
      const pview = new DataView(
          buffer.buffer,
          buffer.byteOffset + base + section.positions.offset,
          section.positions.bytes,
        ),
        list: number[] = [];
      if (section.positions.mode === 0) {
        assert(
          pview.byteLength === section.positions.count * 2,
          'CORRUPT_STORAGE',
          'Position length',
        );
        for (let i = 0; i < section.positions.count; i++) list.push(pview.getUint16(i * 2, true));
      } else if (section.positions.mode === 1) {
        assert(pview.byteLength % 4 === 0, 'CORRUPT_STORAGE', 'Run length');
        for (let i = 0; i < pview.byteLength; i += 4) {
          const start = pview.getUint16(i, true),
            count = pview.getUint16(i + 2, true);
          assert(count > 0 && start + count <= 2048, 'CORRUPT_STORAGE', 'Invalid position run');
          for (let j = 0; j < count; j++) list.push(start + j);
        }
      } else {
        assert(
          section.positions.mode === 2 && pview.byteLength === 256,
          'CORRUPT_STORAGE',
          'Bitmap length',
        );
        for (let i = 0; i < 2048; i++) if (pview.getUint8(i >> 3) & (1 << (i & 7))) list.push(i);
      }
      assert(
        list.length === section.positions.count &&
          list.every((v, i) => v < 2048 && (i === 0 || list[i - 1] < v)),
        'CORRUPT_STORAGE',
        'Invalid page positions',
      );
      bounds(section.offset, section.bytes);
      const vview = new DataView(
          buffer.buffer,
          buffer.byteOffset + base + section.offset,
          section.bytes,
        ),
        count = list.length;
      let value: (index: number) => unknown;
      if (section.name === 'numbers') {
        assert(
          [1, 2, 4, 8].includes(section.type) && vview.byteLength === count * section.type,
          'CORRUPT_STORAGE',
          'Numeric length',
        );
        value = (i) =>
          section.type === 1
            ? vview.getInt8(i)
            : section.type === 2
              ? vview.getInt16(i * 2, true)
              : section.type === 4
                ? vview.getInt32(i * 4, true)
                : vview.getFloat64(i * 8, true);
      } else if (section.name === 'booleans') {
        assert(
          section.type === 0 && vview.byteLength === Math.ceil(count / 8),
          'CORRUPT_STORAGE',
          'Boolean length',
        );
        value = (i) => !!(vview.getUint8(i >> 3) & (1 << (i & 7)));
      } else {
        assert(
          [0, 1, 16, 17].includes(section.type) && vview.byteLength >= 4,
          'CORRUPT_STORAGE',
          'String type',
        );
        const dict = section.type % 16 === 1,
          wide = section.type >= 16,
          n = vview.getUint32(0, true),
          offsets = 4 + (dict ? count * 2 : 0),
          arena = offsets + (n + 1) * 4;
        assert(
          n > 0 && n <= count && (!dict ? n === count : true) && arena <= vview.byteLength,
          'CORRUPT_STORAGE',
          'String directory',
        );
        let prev = 0;
        for (let i = 0; i <= n; i++) {
          const at = vview.getUint32(offsets + i * 4, true);
          assert(
            at >= prev && arena + at <= vview.byteLength && (!wide || at % 2 === 0),
            'CORRUPT_STORAGE',
            'String offset',
          );
          prev = at;
        }
        assert(arena + prev === vview.byteLength, 'CORRUPT_STORAGE', 'Trailing string bytes');
        value = (i) => {
          const index = dict ? vview.getUint16(4 + i * 2, true) : i;
          assert(index < n, 'CORRUPT_STORAGE', 'String dictionary index');
          const start = vview.getUint32(offsets + index * 4, true),
            end = vview.getUint32(offsets + (index + 1) * 4, true);
          if (!wide)
            return decoder.decode(
              new Uint8Array(vview.buffer, vview.byteOffset + arena + start, end - start),
            );
          let text = '';
          for (let p = start; p < end; p += 2)
            text += String.fromCharCode(vview.getUint16(arena + p, true));
          return text;
        };
      }
      this.columns.push({ name: section.name, positions: new Uint16Array(list), value });
      (data as any)[section.name] = { positions: list, values: list.map((_, i) => value(i)) };
    }
    this.extra = decodeFileCells({ row: this.row, column: this.column, data }, this.styles);
    // Validate once, then retain only sparse properties. Scalar objects never live in the page cache.
    for (const cell of Object.values(this.extra)) {
      if (cell.input.type === 'formula') {
        try {
          this.dependencies[key(cell.rowId, cell.columnId)] = parseFormula(cell.input.expression);
        } catch {
          this.dependencies[key(cell.rowId, cell.columnId)] = null;
        }
      }
      if (cell.input.type !== 'blank') delete (cell as Partial<CellRecord>).input;
      if (Object.keys(cell).length === 2) delete this.extra[key(cell.rowId, cell.columnId)];
    }
    this.bytes =
      buffer.byteLength +
      this.columns.reduce((n, c) => n + c.positions.byteLength + 128, 0) +
      JSON.stringify(this.extra).length * 2 +
      Object.keys(this.extra).length * 192 +
      JSON.stringify(this.dependencies).length * 6 +
      1024;
  }
  cell(offset: number): CellRecord | undefined {
    const rowId = this.row * 64 + Math.floor(offset / 32),
      columnId = this.column * 32 + (offset % 32);
    let input: CellInput | undefined;
    for (const col of this.columns) {
      let low = 0,
        high = col.positions.length;
      while (low < high) {
        const mid = (low + high) >>> 1;
        if (col.positions[mid] < offset) low = mid + 1;
        else high = mid;
      }
      if (col.positions[low] !== offset) continue;
      const value = col.value(low);
      input =
        col.name === 'numbers'
          ? { type: 'number', value: value as number }
          : col.name === 'strings'
            ? { type: 'string', value: value as string }
            : col.name === 'booleans'
              ? { type: 'boolean', value: value as boolean }
              : col.name === 'formulas'
                ? { type: 'formula', expression: value as string }
                : { type: 'error', code: value as string };
      break;
    }
    const extra = this.extra[key(rowId, columnId)];
    if (!input && !extra) return;
    return { ...extra, rowId, columnId, input: input ?? { type: 'blank' } };
  }
  get cells(): Record<string, CellRecord> {
    const offsets = new Set<number>(
      Object.values(this.extra).map((c) => (c.rowId % 64) * 32 + (c.columnId % 32)),
    );
    for (const col of this.columns) for (const offset of col.positions) offsets.add(offset);
    return Object.fromEntries(
      [...offsets].map((offset) => {
        const cell = this.cell(offset)!;
        return [key(cell.rowId, cell.columnId), cell];
      }),
    );
  }
}
export const decodeBlock = (encoded: EncodedBlock) => new ColumnPage(encoded);
