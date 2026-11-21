import { type WorkbookSnapshot } from '../packages/core/src/types.js';
import { encodeWorkbookFile, decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import fc from 'fast-check';
import {
  createWorkbook,
  openWorkbook,
  Workbook,
  workbookFileSchema,
  key,
  type WorkbookFile,
} from '@opensheetjs/core';
import { encodeAxis } from '../packages/core/src/file-codec.js';
import { encodeBlock, decodeBlock } from '../packages/core/src/block-codec.js';
import {
  BinaryDirectoryBudget,
  BinaryReader,
  crc32,
  readBinary,
  unpackPage,
  writeBinary,
} from '../packages/core/src/binary-file.js';
import { parseWorkbookJSON } from '../packages/core/src/json-stream.js';
const books: Workbook[] = [];
afterEach(() => books.splice(0).forEach((book) => book.dispose()));
const track = async (snapshot?: WorkbookSnapshot) => {
  const book = snapshot
    ? await new Workbook(encodeWorkbookFile(snapshot)).ready()
    : await createWorkbook();
  books.push(book);
  return book;
};
const streamFile = async (book: Workbook) => {
  let text = '';
  for await (const chunk of book.streamJSON()) text += chunk;
  return JSON.parse(text) as WorkbookFile;
};
const binary = async (book: Workbook) => {
  const parts: Uint8Array[] = [];
  for await (const bytes of book.streamBinary()) parts.push(bytes);
  return new Blob(parts as BlobPart[]);
};
const canonical = (file: WorkbookFile) => ({ ...file, revision: 0, workbookId: 'ignored' });

it('classifies damaged gzip pages and preserves an already-aborted signal', async () => {
  const book = await track();
  await book
    .getSheets()[0]
    .range('A1')
    .setValues([[42]]);
  const damaged = new Uint8Array(await (await binary(book)).arrayBuffer());
  const view = new DataView(damaged.buffer);
  const page = 12 + view.getUint32(4, true),
    length = view.getUint32(page + 16, true);
  view.setUint32(page + 12, 1, true);
  damaged.fill(0, page + 28, page + 28 + length);
  view.setUint32(page + 24, crc32(damaged.subarray(page + 28, page + 28 + length)), true);
  await expect(async () => {
    for await (const _part of readBinary(damaged)) {
      /* consume */
    }
  }).rejects.toMatchObject({ code: 'CORRUPT_STORAGE' });
  await expect((await BinaryReader.open(damaged)).page(0)).rejects.toMatchObject({
    code: 'CORRUPT_STORAGE',
  });
  const controller = new AbortController();
  const reason = new Error('cancelled');
  controller.abort(reason);
  await expect(unpackPage(new Uint8Array([0]), 1, 1, undefined, controller.signal)).rejects.toBe(
    reason,
  );
});

it('bounds binary directory entries at their exact serialized size before retaining them', () => {
  const entry = fc.record({
    sheet: fc.nat(),
    row: fc.nat(),
    column: fc.nat(),
    offset: fc.nat(),
    length: fc.nat(),
  });
  fc.assert(
    fc.property(fc.array(entry, { minLength: 1, maxLength: 100 }), (entries) => {
      const exact = JSON.stringify(entries).length;
      const accepted = new BinaryDirectoryBudget(exact);
      for (const value of entries) accepted.add(value);
      const rejected = new BinaryDirectoryBudget(exact - 1);
      for (const value of entries.slice(0, -1)) rejected.add(value);
      expect(() => rejected.add(entries.at(-1)!)).toThrow(
        expect.objectContaining({ code: 'LIMIT_EXCEEDED' }),
      );
    }),
    { seed: 20261112, numRuns: 300 },
  );
});

it('counts binary directory digit boundaries up to the safe file offset limit', () => {
  const entries = [0, 9, 10, 99, 100, 0xffffffff, Number.MAX_SAFE_INTEGER].map((offset) => ({
    sheet: 0,
    row: 0xffffffff,
    column: 9,
    offset,
    length: 0xffffffff + 28,
  }));
  const size = JSON.stringify(entries).length;
  const accepted = new BinaryDirectoryBudget(size);
  for (const entry of entries) accepted.add(entry);
  const rejected = new BinaryDirectoryBudget(size - 1);
  for (const entry of entries.slice(0, -1)) rejected.add(entry);
  expect(() => rejected.add(entries.at(-1)!)).toThrow(
    expect.objectContaining({ code: 'LIMIT_EXCEEDED' }),
  );
});

it('rejects an over-budget directory before reading the next binary page payload', async () => {
  const book = await track();
  const sheet = book.getSheets()[0];
  await sheet.range('A1').setValues([[1]]);
  await sheet.range('A65').setValues([[2]]);
  const bytes = new Uint8Array(await (await binary(book)).arrayBuffer());
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const first = 12 + view.getUint32(4, true);
  const second = first + 28 + view.getUint32(first + 16, true);
  const chunks = [
    bytes.subarray(0, first),
    bytes.subarray(first, second),
    bytes.subarray(second, second + 28),
    bytes.subarray(second + 28),
  ];
  let next = 0,
    count = 0;
  const source = new ReadableStream<Uint8Array>(
    {
      pull(controller) {
        if (next === chunks.length) controller.close();
        else controller.enqueue(chunks[next++]);
      },
    },
    { highWaterMark: 0 },
  );
  const original = BinaryDirectoryBudget.prototype.add;
  const spy = vi.spyOn(BinaryDirectoryBudget.prototype, 'add').mockImplementation(function (
    this: BinaryDirectoryBudget,
    entry,
  ) {
    if (++count === 2) throw new Error('directory budget');
    return original.call(this, entry);
  });
  try {
    await expect(async () => {
      for await (const _part of readBinary(source)) {
        /* consume */
      }
    }).rejects.toThrow('directory budget');
    expect(next).toBe(3);
    expect(source.locked).toBe(false);
  } finally {
    spy.mockRestore();
  }
});

it('rejects invalid page encoding headers before reading their payload', async () => {
  const book = await track();
  await book
    .getSheets()[0]
    .range('A1')
    .setValues([[42]]);
  const original = new Uint8Array(await (await binary(book)).arrayBuffer());
  const page = 12 + new DataView(original.buffer).getUint32(4, true);
  for (const [codec, decodedLength, code] of [
    [0, 0, 'LIMIT_EXCEEDED'],
    [1, 32 * 1024 * 1024 + 1, 'LIMIT_EXCEEDED'],
    [2, 1, 'INVALID_ARGUMENT'],
    [0, 1, 'INVALID_ARGUMENT'],
  ] as const) {
    const bytes = original.slice();
    const view = new DataView(bytes.buffer);
    view.setUint32(page + 12, codec, true);
    view.setUint32(page + 20, decodedLength, true);
    const chunks = [
      bytes.subarray(0, page),
      bytes.subarray(page, page + 28),
      bytes.subarray(page + 28),
    ];
    let next = 0;
    const source = new ReadableStream<Uint8Array>(
      {
        pull(controller) {
          if (next === chunks.length) controller.close();
          else controller.enqueue(chunks[next++]);
        },
      },
      { highWaterMark: 0 },
    );
    await expect(async () => {
      for await (const _part of readBinary(source)) {
        /* consume */
      }
    }).rejects.toMatchObject({ code });
    expect(next).toBe(2);
    expect(source.locked).toBe(false);
  }
});

it('checks large random-access page headers before loading their payload', async () => {
  const metadata = encodeWorkbookFile(createSnapshot());
  const payload = new Uint8Array(96 * 1024);
  let state = 123456789;
  for (let i = 0; i < payload.length; i++) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    payload[i] = state & 255;
  }
  async function* pages() {
    yield { sheet: 0, row: 0, column: 0, bytes: payload };
  }
  const chunks: Uint8Array[] = [];
  for await (const chunk of writeBinary(metadata, pages())) chunks.push(chunk);
  const original = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
  const offset = 12 + new DataView(original.buffer).getUint32(4, true);
  expect(new DataView(original.buffer).getUint32(offset + 16, true)).toBeGreaterThan(64 * 1024);
  const blob = new Blob([original]);
  const reader = await BinaryReader.open(blob);
  const slices = vi.spyOn(blob, 'slice');
  expect(await reader.page(0)).toEqual(payload);
  expect(slices.mock.calls.map(([start, end]) => [start, end])).toEqual([
    [offset, offset + 28],
    [offset + 28, offset + reader.directory[0].length],
  ]);
  slices.mockRestore();
  const corrupt = original.slice();
  corrupt[offset + 28] ^= 1;
  await expect((await BinaryReader.open(corrupt)).page(0)).rejects.toMatchObject({
    code: 'INVALID_ARGUMENT',
  });

  if (typeof CompressionStream !== 'undefined' && typeof DecompressionStream !== 'undefined') {
    const repeated = new Uint8Array(payload.length * 2);
    for (let i = 0; i < payload.length; i += 16384) {
      repeated.set(payload.subarray(i, i + 16384), i * 2);
      repeated.set(payload.subarray(i, i + 16384), i * 2 + 16384);
    }
    async function* compressedPages() {
      yield { sheet: 0, row: 0, column: 0, bytes: repeated };
    }
    const compressedChunks: Uint8Array[] = [];
    for await (const chunk of writeBinary(metadata, compressedPages()))
      compressedChunks.push(chunk);
    const compressed = new Uint8Array(await new Blob(compressedChunks as BlobPart[]).arrayBuffer());
    const view = new DataView(compressed.buffer);
    const compressedOffset = 12 + view.getUint32(4, true);
    expect(view.getUint32(compressedOffset + 12, true)).toBe(1);
    expect(view.getUint32(compressedOffset + 16, true)).toBeGreaterThan(64 * 1024);
    expect(await (await BinaryReader.open(compressed)).page(0)).toEqual(repeated);
  }

  for (const [codec, length, code] of [
    [0, 0, 'LIMIT_EXCEEDED'],
    [1, 32 * 1024 * 1024 + 1, 'LIMIT_EXCEEDED'],
    [2, 1, 'INVALID_ARGUMENT'],
    [0, 1, 'INVALID_ARGUMENT'],
  ] as const) {
    const damaged = original.slice();
    const view = new DataView(damaged.buffer);
    view.setUint32(offset + 12, codec, true);
    view.setUint32(offset + 20, length, true);
    const source = new Blob([damaged]);
    const random = await BinaryReader.open(source);
    const spy = vi.spyOn(source, 'slice');
    await expect(random.page(0)).rejects.toMatchObject({ code });
    expect(spy.mock.calls.map(([start, end]) => [start, end])).toEqual([[offset, offset + 28]]);
    spy.mockRestore();
  }
});

it('accepts empty binary stream chunks while still rejecting trailing bytes and invalid chunks', async () => {
  const book = await track();
  await book
    .getSheets()[0]
    .range('A1')
    .setValues([[42]]);
  const bytes = new Uint8Array(await (await binary(book)).arrayBuffer());
  const stream = (trailing: unknown[] = []) =>
    new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array());
        controller.enqueue(bytes.subarray(0, 8));
        controller.enqueue(new Uint8Array());
        controller.enqueue(bytes.subarray(8));
        controller.enqueue(new Uint8Array());
        for (const item of trailing) controller.enqueue(item as Uint8Array);
        controller.close();
      },
    });
  const parts = [];
  for await (const part of readBinary(stream())) parts.push(part);
  expect(parts.some((part) => 'metadata' in part)).toBe(true);
  expect(parts.some((part) => 'bytes' in part)).toBe(true);
  for (const [item, message] of [
    [new Uint8Array([1]), 'Trailing binary data'],
    ['bad', 'Uint8Array'],
  ] as const) {
    await expect(async () => {
      for await (const _part of readBinary(stream([item]))) {
        /* consume */
      }
    }).rejects.toMatchObject({
      code: 'INVALID_ARGUMENT',
      message: expect.stringContaining(message),
    });
  }
});

it('allows timer-driven cancellation of a binary stream producing endless empty chunks', async () => {
  const controller = new AbortController();
  let cancelled = false;
  const source = new ReadableStream<Uint8Array>({
    pull(stream) {
      stream.enqueue(new Uint8Array());
    },
    cancel() {
      cancelled = true;
    },
  });
  const timer = setTimeout(() => controller.abort(), 0);
  try {
    await expect(async () => {
      for await (const _part of readBinary(source, { signal: controller.signal })) {
        /* consume */
      }
    }).rejects.toMatchObject({ name: 'AbortError' });
    expect(cancelled).toBe(true);
    expect(source.locked).toBe(false);
  } finally {
    clearTimeout(timer);
  }
});

it('cancels a binary section fed by endless one-byte chunks and still reads finite tiny chunks', async () => {
  const header = new Uint8Array(12);
  header.set([79, 83, 66, 3]);
  new DataView(header.buffer).setUint32(4, 32 * 1024 * 1024, true);
  const abort = new AbortController();
  let first = true,
    cancelled = false;
  const input = new ReadableStream<Uint8Array>({
    pull(controller) {
      controller.enqueue(first ? header : new Uint8Array([32]));
      first = false;
    },
    cancel() {
      cancelled = true;
    },
  });
  const timer = setTimeout(() => abort.abort(), 0);
  try {
    await expect(async () => {
      for await (const _part of readBinary(input, { signal: abort.signal })) {
        /* consume */
      }
    }).rejects.toMatchObject({ name: 'AbortError' });
    expect(cancelled).toBe(true);
    expect(input.locked).toBe(false);
  } finally {
    clearTimeout(timer);
  }
  const bytes = new Uint8Array(await (await binary(await track())).arrayBuffer());
  let offset = 0;
  const finite = new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset === bytes.length) controller.close();
      else controller.enqueue(bytes.subarray(offset, ++offset));
    },
  });
  const parts = [];
  for await (const part of readBinary(finite)) parts.push(part);
  expect(parts.some((part) => 'metadata' in part)).toBe(true);
});

it('rejects malformed JSON stream bytes and permits cancellation of nonproductive streams', async () => {
  for (const chunks of [[new Uint8Array([0xff])], ['bad'], [undefined]]) {
    const input = new ReadableStream<Uint8Array>({
      start(controller) {
        for (const chunk of chunks) controller.enqueue(chunk as Uint8Array);
        controller.close();
      },
    });
    await expect(async () => {
      for await (const _part of parseWorkbookJSON(input)) {
        /* consume */
      }
    }).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  }
  for (const chunk of [
    new Uint8Array(),
    new TextEncoder().encode('   '),
    new TextEncoder().encode('a'),
  ]) {
    const abort = new AbortController();
    let cancelled = false;
    const input = new ReadableStream<Uint8Array>({
      pull(controller) {
        controller.enqueue(chunk);
      },
      cancel() {
        cancelled = true;
      },
    });
    const timer = setTimeout(() => abort.abort(), 0);
    try {
      await expect(async () => {
        for await (const _part of parseWorkbookJSON(input, { signal: abort.signal })) {
          /* consume */
        }
      }).rejects.toMatchObject({ name: 'AbortError' });
      expect(cancelled).toBe(true);
      expect(input.locked).toBe(false);
    } finally {
      clearTimeout(timer);
    }
  }
});

it('allows cancellation during productive JSON streams and a single large whitespace chunk', async () => {
  for (const largeChunk of [false, true]) {
    const abort = new AbortController();
    let first = true,
      cancelled = false;
    const encoder = new TextEncoder();
    const input = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (largeChunk) {
          if (first) controller.enqueue(encoder.encode(' '.repeat(8 * 1024 * 1024)));
          else return;
        } else controller.enqueue(encoder.encode(first ? '[0,' : '0,'));
        first = false;
      },
      cancel() {
        cancelled = true;
      },
    });
    const timer = setTimeout(() => abort.abort(), 0);
    try {
      await expect(async () => {
        for await (const _part of parseWorkbookJSON(input, { signal: abort.signal })) {
          /* consume */
        }
      }).rejects.toMatchObject({ name: 'AbortError' });
      expect(cancelled).toBe(true);
      expect(input.locked).toBe(false);
    } finally {
      clearTimeout(timer);
    }
  }
});

it('decodes large JSON chunks in bounded slices across UTF-8 boundaries and rejects long tokens early', async () => {
  const prefix = '{"value":"';
  const valid = new TextEncoder().encode(' '.repeat(65535 - prefix.length) + prefix + '😀"}');
  const spy = vi.spyOn(TextDecoder.prototype, 'decode');
  const input = (bytes: Uint8Array) =>
    new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(bytes);
        controller.close();
      },
    });
  try {
    let metadata: unknown;
    for await (const part of parseWorkbookJSON(input(valid)))
      if ('metadata' in part) metadata = part.metadata;
    expect(metadata).toEqual({ value: '😀' });
    const sizes = spy.mock.calls.map(([bytes]) => bytes?.byteLength ?? 0);
    expect(Math.max(...sizes)).toBeLessThanOrEqual(65536);
    spy.mockClear();
    const oversized = new TextEncoder().encode('{"value":"' + 'x'.repeat(2 * 1024 * 1024) + '"}');
    await expect(async () => {
      for await (const _part of parseWorkbookJSON(input(oversized))) {
        /* consume */
      }
    }).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const totalDecoded = spy.mock.calls.reduce((sum, [bytes]) => sum + (bytes?.byteLength ?? 0), 0);
    expect(totalDecoded).toBeLessThan(oversized.byteLength);
    expect(spy.mock.calls.every(([bytes]) => (bytes?.byteLength ?? 0) <= 65536)).toBe(true);
  } finally {
    spy.mockRestore();
  }
});

describe('native column files and binary pages', () => {
  it('returns independently editable metadata without mutating the live workbook', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 10, columns: 3 }] });
    snapshot.extensions.fixture = { nested: { label: 'original' } };
    snapshot.styles.custom = { bold: true };
    snapshot.sheets[0].rows['0'] = { size: 42 };
    snapshot.sheets[0].cells['0:0'] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'number', value: 7 },
      styleId: 'custom',
    };
    const book = await track(snapshot);
    const revision = book.revision;
    const file = await book.toJSON();
    (file.extensions.fixture as { nested: { label: string } }).nested.label = 'changed';
    file.sheets[0].name = 'Edited';
    file.sheets[0].rows.meta!['0'].size = 55;
    file.styles[0].bold = false;
    expect(book.revision).toBe(revision);
    const unchanged = await book.toJSON();
    expect(unchanged.extensions.fixture).toEqual({ nested: { label: 'original' } });
    expect(unchanged.sheets[0].name).toBe('Data');
    expect(unchanged.sheets[0].rows.meta!['0'].size).toBe(42);
    expect(unchanged.styles[0].bold).toBe(true);
    await book.importJSON(file);
    expect((await book.toJSON()).extensions.fixture).toEqual({ nested: { label: 'changed' } });
    expect(book.getSheets()[0].name).toBe('Edited');
  });
  it('preserves every type, sparse property, negative zero and Unicode through JSON and binary', async () => {
    const snapshot = createSnapshot({
      sheets: [{ name: 'One', rows: 130, columns: 40 }, { name: 'Two' }],
    });
    snapshot.sheetOrder.reverse();
    snapshot.dateSystem = '1904';
    snapshot.extensions = { test: { value: 0 } };
    snapshot.styles.s = { bold: true, numberFormat: '0%' };
    const sheet = snapshot.sheets[0];
    sheet.rowOrder.splice(5, 0, sheet.nextRowId++);
    sheet.rowOrder.reverse();
    sheet.rows[sheet.rowOrder[9]] = { hidden: true, size: 31 };
    sheet.columns[2] = { size: 111 };
    sheet.freeze = { rows: 1, columns: 2 };
    const inputs = [
      { type: 'blank' },
      { type: 'string', value: '' },
      { type: 'string', value: '=A1\ud800你好😀' },
      { type: 'number', value: -0 },
      { type: 'boolean', value: false },
      { type: 'formula', expression: 'SUM(A1:A2)' },
      { type: 'error', code: '#REF!' },
    ] as const;
    inputs.forEach((input, i) => {
      const rowId = sheet.rowOrder[i * 18],
        columnId = sheet.columnOrder[i * 5];
      sheet.cells[key(rowId, columnId)] = {
        rowId,
        columnId,
        input,
        styleId: 's',
        note: '注释',
        link: { target: 'https://example.com', tooltip: 'link' },
        numberFormat: '0.00',
        cached: { type: 'number', value: -0 },
      };
    });
    const file = encodeWorkbookFile(snapshot);
    expect(workbookFileSchema.safeParse(file).success).toBe(true);
    expect(encodeWorkbookFile(decodeWorkbookFile(JSON.parse(JSON.stringify(file))))).toEqual(file);
    const book = await track(snapshot);
    expect(await streamFile(book)).toEqual(await book.toJSON());
    const backup = await binary(book);
    const reader = await BinaryReader.open(backup);
    for (const index of reader.directory.keys()) {
      const page = decodeBlock(await reader.page(index));
      const records = [...page.records()];
      expect(records).toHaveLength(Object.keys(page.cells).length);
      expect(
        Object.fromEntries(records.map((cell) => [key(cell.rowId, cell.columnId), cell])),
      ).toEqual(page.cells);
    }
    const restored = await track();
    await restored.importBinary(backup);
    expect(canonical(await restored.toJSON())).toEqual(canonical(await book.toJSON()));
    const cell = await book.getCell(
      sheet.id,
      sheet.rowOrder.indexOf(sheet.rowOrder[3 * 18]),
      sheet.columnOrder.indexOf(sheet.columnOrder[3 * 5]),
    );
    expect(Object.is(cell!.input.type === 'number' ? cell!.input.value : 1, -0)).toBe(true);
  });
  it('distinguishes unsupported native versions from malformed data without changing the workbook', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const revision = book.revision;
    const file = await book.toJSON();
    for (const version of [2, 4]) {
      const incompatible = { ...file, schemaVersion: version };
      expect(() => decodeWorkbookFile(incompatible)).toThrow(
        expect.objectContaining({ code: 'UNSUPPORTED_VERSION' }),
      );
      await expect(book.importJSON(JSON.stringify(incompatible))).rejects.toMatchObject({
        code: 'UNSUPPORTED_VERSION',
      });
      expect(book.revision).toBe(revision);
      expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
    }
    await expect(book.importJSON('{"schemaVersion":99}')).rejects.toMatchObject({
      code: 'INVALID_ARGUMENT',
    });
    const bytes = new Uint8Array(await (await binary(book)).arrayBuffer());
    bytes[3] = 4;
    await expect(BinaryReader.open(bytes)).rejects.toMatchObject({ code: 'UNSUPPORTED_VERSION' });
    await expect(book.importBinary(bytes)).rejects.toMatchObject({ code: 'UNSUPPORTED_VERSION' });
    expect(book.revision).toBe(revision);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
    expect(await book.undo()).toBe(true);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
    bytes[0] = 0;
    await expect(BinaryReader.open(bytes)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  });
  it('classifies invalid file metadata and directory text consistently', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const revision = book.revision;
    const original = new Uint8Array(await (await binary(book)).arrayBuffer());
    const originalView = new DataView(original.buffer);
    const metadataLength = originalView.getUint32(4, true);
    const footerOffset = original.length - 24;
    const directoryOffset = Number(originalView.getBigUint64(footerOffset + 4, true));
    const directoryLength = Number(originalView.getBigUint64(footerOffset + 12, true));
    for (const target of ['metadata', 'directory'] as const) {
      for (const invalid of [0x21, 0xff]) {
        const damaged = original.slice();
        const view = new DataView(damaged.buffer);
        const start = target === 'metadata' ? 12 : directoryOffset;
        const length = target === 'metadata' ? metadataLength : directoryLength;
        damaged[start] = invalid;
        view.setUint32(
          target === 'metadata' ? 8 : footerOffset + 20,
          crc32(damaged.subarray(start, start + length)),
          true,
        );
        await expect(BinaryReader.open(damaged)).rejects.toMatchObject({
          code: 'INVALID_ARGUMENT',
        });
        await expect(
          (async () => {
            for await (const _ of readBinary(damaged)) {
              /* consume */
            }
          })(),
        ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
        await expect(book.importBinary(damaged)).rejects.toMatchObject({
          code: 'INVALID_ARGUMENT',
        });
        expect(book.revision).toBe(revision);
        expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
      }
    }
    const damaged = original.slice();
    damaged.fill(0x20, directoryOffset, directoryOffset + directoryLength);
    damaged.set(new TextEncoder().encode('[null]'), directoryOffset);
    new DataView(damaged.buffer).setUint32(
      footerOffset + 20,
      crc32(damaged.subarray(directoryOffset, directoryOffset + directoryLength)),
      true,
    );
    await expect(BinaryReader.open(damaged)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    await expect(book.importBinary(damaged)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    expect(book.revision).toBe(revision);
  });
  it('rejects a damaged directory marker in both binary readers without changing a workbook', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const revision = book.revision;
    const original = new Uint8Array(await (await binary(book)).arrayBuffer());
    const footer = new DataView(original.buffer, original.length - 24, 24);
    const directory = Number(footer.getBigUint64(4, true));
    await BinaryReader.open(original);
    for (const offset of [directory - 8, directory - 4]) {
      const damaged = original.slice();
      damaged[offset] ^= 1;
      await expect(BinaryReader.open(damaged)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      await expect(
        (async () => {
          for await (const _part of readBinary(damaged)) {
            /* Consume the complete file. */
          }
        })(),
      ).rejects.toThrow();
      await expect(book.importBinary(damaged)).rejects.toThrow();
      expect(book.revision).toBe(revision);
      expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
    }
    expect(await book.undo()).toBe(true);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
  });
  it('rejects duplicate or unknown-sheet binary pages before staging them', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const revision = book.revision;
    const source = await BinaryReader.open(await binary(book));
    expect(source.directory).toHaveLength(1);
    const bytes = await source.page(0);
    for (const mode of ['duplicate', 'sheet'] as const) {
      async function* pages() {
        yield { sheet: 0, row: 0, column: 0, bytes };
        yield { sheet: mode === 'sheet' ? 1 : 0, row: 0, column: 0, bytes };
      }
      const parts: Uint8Array[] = [];
      for await (const part of writeBinary(source.metadata, pages())) parts.push(part);
      const malformed = new Blob(parts as BlobPart[]);
      await expect(BinaryReader.open(malformed)).rejects.toMatchObject({
        code: 'INVALID_ARGUMENT',
      });
      let pagesRead = 0;
      await expect(
        (async () => {
          for await (const item of readBinary(malformed)) if ('bytes' in item) pagesRead++;
        })(),
      ).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      expect(pagesRead).toBe(1);
      await expect(book.importBinary(malformed)).rejects.toMatchObject({
        code: 'INVALID_ARGUMENT',
      });
      expect(book.revision).toBe(revision);
      expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
    }
    expect(await book.undo()).toBe(true);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[null]]);
  });
  it('rejects binary cells outside sheet axes or with missing styles after staging', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
    const book = await track(snapshot);
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[91]]);
    const revision = book.revision;
    for (const kind of ['axis', 'style'] as const) {
      const rowId = kind === 'axis' ? 2 : 0;
      const cell = {
        rowId,
        columnId: 0,
        input: { type: 'number' as const, value: 42 },
        ...(kind === 'style' ? { styleId: 's0' } : {}),
      };
      async function* pages() {
        yield {
          sheet: 0,
          row: 0,
          column: 0,
          bytes: encodeBlock({ [key(rowId, 0)]: cell }, [], [], {}),
        };
      }
      const parts: Uint8Array[] = [];
      for await (const part of writeBinary(encodeWorkbookFile(snapshot), pages())) parts.push(part);
      await expect(book.importBinary(new Blob(parts as BlobPart[]))).rejects.toThrow(
        kind === 'axis' ? 'Cell outside sheet axes' : 'Unknown style',
      );
      expect(book.revision).toBe(revision);
      expect(await sheet.range('A1').getValues()).toEqual([[91]]);
    }
  });
  it('rejects unsupported binary metadata before yielding or staging any page', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const revision = book.revision;
    const source = await BinaryReader.open(await binary(book));
    const bytes = await source.page(0);
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes };
    }
    const parts: Uint8Array[] = [];
    const future = { ...source.metadata, schemaVersion: 4 } as unknown as WorkbookFile;
    for await (const part of writeBinary(future, pages())) parts.push(part);
    const malformed = new Blob(parts as BlobPart[]);
    let yielded = 0;
    await expect(
      (async () => {
        for await (const _item of readBinary(malformed)) yielded++;
      })(),
    ).rejects.toMatchObject({ code: 'UNSUPPORTED_VERSION' });
    expect(yielded).toBe(0);
    await expect(book.importBinary(malformed)).rejects.toMatchObject({
      code: 'UNSUPPORTED_VERSION',
    });
    expect(book.revision).toBe(revision);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
  });
  it('revalidates mutable directory entries before reading a random-access page', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[1]]);
    const reader = await BinaryReader.open(await binary(book));
    const original = { ...reader.directory[0] };
    expect(await reader.page(0)).toBeInstanceOf(Uint8Array);
    for (const update of [
      { offset: -1 },
      { length: 12 },
      { length: Number.MAX_SAFE_INTEGER },
      { offset: Number.MAX_SAFE_INTEGER },
    ]) {
      Object.assign(reader.directory[0], original, update);
      await expect(reader.page(0)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    }
    Object.assign(reader.directory[0], original);
    expect(await reader.page(0)).toBeInstanceOf(Uint8Array);
    await expect(reader.page(-1)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  });
  it('rejects oversized random-access page frames before materializing them', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[1]]);
    const original = new Uint8Array(await (await binary(book)).arrayBuffer());
    const footer = new DataView(original.buffer, original.length - 24, 24);
    const directoryOffset = Number(footer.getBigUint64(4, true));
    const directoryLength = Number(footer.getBigUint64(12, true));
    const entries = JSON.parse(
      new TextDecoder().decode(
        original.subarray(directoryOffset, directoryOffset + directoryLength),
      ),
    );
    expect(entries).toHaveLength(1);
    const entry = entries[0];
    const payload = new Uint8Array(32 * 1024 * 1024 + 1);
    const makeFile = (size: number) => {
      entry.length = 28 + size;
      const directory = new TextEncoder().encode(JSON.stringify(entries));
      const marker = new Uint8Array(8),
        markerView = new DataView(marker.buffer);
      markerView.setUint32(0, 0xffffffff, true);
      markerView.setUint32(4, directory.length, true);
      const trailer = new Uint8Array(24),
        trailerView = new DataView(trailer.buffer);
      trailer.set([79, 83, 66, 69]);
      trailerView.setBigUint64(4, BigInt(entry.offset + entry.length + 8), true);
      trailerView.setBigUint64(12, BigInt(directory.length), true);
      trailerView.setUint32(20, crc32(directory), true);
      const frame = new Uint8Array(28),
        frameView = new DataView(frame.buffer);
      frameView.setUint32(16, size, true);
      frameView.setUint32(20, size, true);
      return new Blob([
        original.subarray(0, entry.offset),
        frame,
        payload.subarray(0, size),
        marker,
        directory,
        trailer,
      ] as BlobPart[]);
    };
    // Opening validates directory bounds; this synthetic frame is not a valid page payload.
    const atLimit = await BinaryReader.open(makeFile(32 * 1024 * 1024));
    atLimit.directory[0].length++;
    await expect(atLimit.page(0)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    const forged = makeFile(payload.length);
    await expect(BinaryReader.open(forged)).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
    await expect(
      (async () => {
        for await (const _part of readBinary(forged)) {
          /* Consume until rejection. */
        }
      })(),
    ).rejects.toMatchObject({ code: 'LIMIT_EXCEEDED' });
  });
  it('rejects unknown nested native fields without replacing existing data or history', async () => {
    const book = await track();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[91]]);
    await sheet.range('A1').setStyle({ bold: true });
    const file = await book.toJSON();
    const before = await book.toJSON();
    const mutations: Array<(file: any) => void> = [
      (file) => {
        file.styles[0].unknown = true;
      },
      (file) => {
        file.sheets[0].freeze.unknown = true;
      },
      (file) => {
        file.sheets[0].rows.meta = { 0: { size: 20, unknown: true } };
      },
      (file) => {
        file.sheets[0].columns.meta = { 0: { hidden: true, unknown: true } };
      },
      (file) => {
        file.sheets[0].blocks[0].data.cached = {
          positions: [0],
          values: [{ type: 'number', value: 91, unknown: true }],
        };
      },
    ];
    for (const mutate of mutations) {
      const invalid = structuredClone(file);
      mutate(invalid);
      expect(workbookFileSchema.safeParse(invalid).success).toBe(false);
      await expect(book.importJSON(JSON.stringify(invalid))).rejects.toMatchObject({
        code: 'INVALID_ARGUMENT',
      });
      expect(await book.toJSON()).toEqual(before);
    }
    expect(await book.undo()).toBe(true);
    expect(await sheet.range('A1').getValues()).toEqual([[91]]);
  });
  it('round-trips arbitrary order and compresses forward/reverse segments', () => {
    fc.assert(
      fc.property(
        fc.uniqueArray(fc.integer({ min: 0, max: 500 }), { minLength: 1, maxLength: 100 }),
        (order) => {
          const snapshot = createSnapshot();
          snapshot.sheets[0].rowOrder = order;
          snapshot.sheets[0].nextRowId = 501;
          const file = encodeWorkbookFile(snapshot);
          expect(encodeWorkbookFile(decodeWorkbookFile(file))).toEqual(file);
        },
      ),
    );
    expect(encodeAxis([3, 2, 1, 5, 6], 7).order).toEqual([
      [3, 3, -1],
      [5, 2, 1],
    ]);
  });
  it('rejects malformed lengths, duplicate cells/pages, old schemas and unsafe objects', () => {
    const snapshot = createSnapshot();
    snapshot.sheets[0].cells[key(0, 0)] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'number', value: 0 },
    };
    const file = encodeWorkbookFile(snapshot);
    const mutations: Array<(f: any) => void> = [
      (f) => (f.schemaVersion = 2),
      (f) => (f.sheets[0].rows.order = [[0, 100001, 1]]),
      (f) => f.sheets[0].blocks.push(f.sheets[0].blocks[0]),
      (f) => f.sheets[0].blocks[0].data.numbers.positions.push(0),
      (f) => (f.sheets[0].blocks[0].data.numbers.values[0] = 'bad'),
      (f) => (f.sheets[0].blocks[0].data.styles = [[0, 1, 99]]),
      (f) => (f.sheets[0].blocks[0].row = 100),
      (f) => (f.sheets[0].rows.meta = { '1000': {} }),
      (f) => f.sheets.push(f.sheets[0]),
    ];
    for (const mutate of mutations) {
      const copy = structuredClone(file);
      mutate(copy);
      expect(() => decodeWorkbookFile(copy)).toThrow();
    }
    expect(() => decodeWorkbookFile(JSON.parse('{"__proto__":{}}'))).toThrow();
  });
  it('exports sparse distant cells by occupied pages and expires stale imported caches', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Sparse', rows: 100000, columns: 1000 }] });
    const sh = snapshot.sheets[0];
    sh.cells[key(0, 0)] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'formula', expression: '1+1' },
      cached: { type: 'number', value: 2 },
    };
    sh.cells[key(99999, 999)] = {
      rowId: 99999,
      columnId: 999,
      input: { type: 'boolean', value: false },
    };
    const book = await track(snapshot),
      read = vi.spyOn(book, 'streamRange'),
      file = await streamFile(book);
    expect(read).not.toHaveBeenCalled();
    expect(file.sheets[0].blocks).toHaveLength(2);
    expect(JSON.stringify(file).length).toBeLessThan(1700);
    await book
      .getSheets()[0]
      .range('B1')
      .setValues([[1]]);
    expect(
      decodeWorkbookFile(await book.toJSON()).sheets[0].cells[key(0, 0)].cached,
    ).toBeUndefined();
  });
  it('imports chunks with reordered metadata and preserves the original on a late error', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[7]]);
    const file = await book.toJSON();
    const reordered = {
      sheets: file.sheets,
      styles: file.styles,
      ...Object.fromEntries(
        Object.entries(file).filter(([key]) => key !== 'sheets' && key !== 'styles'),
      ),
    };
    const text = JSON.stringify(reordered);
    const bytes = new TextEncoder().encode(text);
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        for (let i = 0; i < bytes.length; i += 3) controller.enqueue(bytes.slice(i, i + 3));
        controller.close();
      },
    });
    await book.importJSON(stream);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[7]]);
    const before = await book.toJSON();
    await expect(book.importJSON(text.slice(0, -1))).rejects.toThrow();
    expect(await book.toJSON()).toEqual(before);
    await expect(
      book.importJSON(text.replace('"schemaVersion":3', '"schemaVersion":3,"schemaVersion":3')),
    ).rejects.toThrow(/Duplicate/);
    for (const text of ['{"x":1,}', '{"x":[1,]}', '{}{}', '{"x":01}']) {
      await expect(async () => {
        for await (const _ of parseWorkbookJSON(text)) {
        }
      }).rejects.toThrow();
    }
  });
  it('allows independent binary page reads and rejects corrupted input atomically', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1:B1')
      .setValues([[0, false]]);
    const blob = await binary(book);
    const reader = await BinaryReader.open(blob);
    expect(Object.values(decodeBlock(await reader.page(0)).cells).map((c) => c.input)).toEqual([
      { type: 'number', value: 0 },
      { type: 'boolean', value: false },
    ]);
    const bytes = new Uint8Array(await blob.arrayBuffer());
    bytes[reader.directory[0].offset + 28] ^= 1;
    const before = await book.toJSON();
    await expect(book.importBinary(bytes)).rejects.toThrow(/checksum/i);
    expect(await book.toJSON()).toEqual(before);
  });
  it('preserves identities through insertion, undo and redo without rewriting unrelated data pages', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1:B2')
      .setValues([
        [0, false],
        ['', 'x'],
      ]);
    const before = await book.toJSON();
    await book.getSheets()[0].insertRows(0);
    await book.getSheets()[0].insertColumns(1);
    const after = await book.toJSON();
    await book.undo();
    await book.undo();
    expect((await book.toJSON()).sheets).toEqual(before.sheets);
    await book.redo();
    await book.redo();
    expect((await book.toJSON()).sheets).toEqual(after.sheets);
  });
  it('aborts exports and rejects mixed revisions', async () => {
    const book = await track();
    const abort = new AbortController();
    abort.abort();
    await expect(book.streamJSON({ signal: abort.signal }).next()).rejects.toMatchObject({
      code: 'ABORTED',
    });
    const stream = book.streamJSON();
    await stream.next();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[1]]);
    await expect(stream.next()).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
  });
  it('retains runtime style identities through metadata persistence and reopening', async () => {
    const database = `styles-${crypto.randomUUID()}`;
    const book = await createWorkbook({}, { database });
    books.push(book);
    const sheetId = book.getSheets()[0].id;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[0.25]]);
    await book.getSheets()[0].range('A1').setStyle({ bold: true, numberFormat: '0%' });
    const before = await book.toJSON();
    await book.close();
    const reopened = await openWorkbook(book.id, { database });
    books.push(reopened);
    const cell = await reopened.getCell(sheetId, 0, 0);
    expect(reopened.getStyle(cell?.styleId)).toEqual({ bold: true, numberFormat: '0%' });
    expect(await reopened.toJSON()).toEqual(before);
    expect(await streamFile(reopened)).toEqual(before);
    const restored = await track();
    await restored.importBinary(await binary(reopened));
    expect(canonical(await restored.toJSON())).toEqual(canonical(before));
  });
  it('omits unused styles consistently across complete JSON, streaming JSON and binary', async () => {
    const snapshot = createSnapshot();
    snapshot.styles.unused = { italic: true };
    snapshot.styles.used = { bold: true };
    const sh = snapshot.sheets[0];
    sh.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'blank' }, styleId: 'used' };
    const book = await track(snapshot);
    const file = await book.toJSON();
    expect(file.styles).toEqual([{ bold: true }]);
    expect(await streamFile(book)).toEqual(file);
    const restored = await track();
    await restored.importBinary(await binary(book));
    expect(canonical(await restored.toJSON())).toEqual(canonical(file));
    await book.getSheets()[0].range('A1').clear({ all: true });
    expect((await book.toJSON()).styles).toEqual([]);
    await book.undo();
    expect((await book.toJSON()).styles).toEqual([{ bold: true }]);
  });
  it('keeps repeated styles and text compact in binary pages', () => {
    const cells: Record<string, import('@opensheetjs/core').CellRecord> = {};
    for (let i = 0; i < 64; i++)
      cells[key(i, 0)] = {
        rowId: i,
        columnId: 0,
        input: { type: 'string', value: 'repeated' },
        styleId: 'style_deadbeef',
      };
    const encoded = encodeBlock(cells, [], [], {});
    expect(encoded).toBeInstanceOf(Uint8Array);
    expect(encoded.length).toBeLessThan(1500);
    expect(decodeBlock(encoded).cells).toEqual(cells);
  });
  it.each([
    ['\uFEFFfirst', '\uFEFFsecond', '\uFEFF'],
    Array(16).fill('\uFEFFrepeat'),
    ['\uFEFFwide', '\uD800'],
  ])('preserves leading U+FEFF in encoded string columns: %j', (...values: string[]) => {
    const cells: Record<string, import('@opensheetjs/core').CellRecord> = {};
    values.forEach((value, row) => {
      cells[key(row, 0)] = { rowId: row, columnId: 0, input: { type: 'string', value } };
    });
    expect(decodeBlock(encodeBlock(cells, [], [], {})).cells).toEqual(cells);
  });
  it('preserves scalar cells, formulas and sparse properties in a mixed binary page', () => {
    const cells: Record<string, import('@opensheetjs/core').CellRecord> = {
      '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: -0 } },
      '0:1': { rowId: 0, columnId: 1, input: { type: 'string', value: 'text' }, note: '' },
      '0:2': { rowId: 0, columnId: 2, input: { type: 'boolean', value: false } },
      '0:3': { rowId: 0, columnId: 3, input: { type: 'formula', expression: 'A1+1' } },
      '0:4': { rowId: 0, columnId: 4, input: { type: 'blank' }, styleId: 's' },
      '0:5': {
        rowId: 0,
        columnId: 5,
        input: { type: 'error', code: '#REF!' },
        cached: { type: 'string', value: '' },
      },
    };
    const page = decodeBlock(encodeBlock(cells, [], [], {}));
    expect(page.cells).toEqual(cells);
    expect(
      Object.fromEntries([...page.records()].map((cell) => [key(cell.rowId, cell.columnId), cell])),
    ).toEqual(cells);
    expect(page.dependencies['0:3']).toBeTruthy();
    expect(Object.is((page.cell(0)!.input as { value: number }).value, -0)).toBe(true);
  });
  it('preserves valid inline scalar columns in native pages', () => {
    const bytes = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } } },
      [],
      [],
      {},
    );
    const length = new DataView(bytes.buffer, bytes.byteOffset).getUint32(4, true);
    const meta = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + length)));
    meta.sections = [];
    meta.data.numbers = { positions: [0], values: [4] };
    meta.data.strings = { positions: [1], values: ['hello'] };
    meta.data.booleans = { positions: [2], values: [false] };
    meta.data.formulas = { positions: [3], values: ['A1+1'] };
    meta.data.errors = { positions: [4], values: ['#REF!'] };
    const header = new TextEncoder().encode(JSON.stringify(meta));
    const base = Math.ceil((8 + header.length) / 8) * 8;
    const inline = new Uint8Array(base);
    inline.set(bytes.subarray(0, 4));
    new DataView(inline.buffer).setUint32(4, header.length, true);
    inline.set(header, 8);
    const page = decodeBlock(inline);
    const expected = [
      { type: 'number', value: 4 },
      { type: 'string', value: 'hello' },
      { type: 'boolean', value: false },
      { type: 'formula', expression: 'A1+1' },
      { type: 'error', code: '#REF!' },
    ];
    expect(expected.map((_, index) => page.cell(index)?.input)).toEqual(expected);
    expect([...page.records()].map((cell) => cell.input)).toEqual(expected);
    expect(page.dependencies['0:3']).toBeTruthy();
    const mixedBytes = encodeBlock(
      {
        '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } },
        '0:1': { rowId: 0, columnId: 1, input: { type: 'string', value: 'hello' } },
      },
      [],
      [],
      {},
    );
    const mixedLength = new DataView(mixedBytes.buffer, mixedBytes.byteOffset).getUint32(4, true);
    const oldBase = Math.ceil((8 + mixedLength) / 8) * 8;
    const mixedMeta = JSON.parse(new TextDecoder().decode(mixedBytes.subarray(8, 8 + mixedLength)));
    mixedMeta.sections = mixedMeta.sections.filter(
      (section: { name: string }) => section.name !== 'numbers',
    );
    mixedMeta.data.numbers = { positions: [0], values: [4] };
    const mixedHeader = new TextEncoder().encode(JSON.stringify(mixedMeta));
    const mixedBase = Math.ceil((8 + mixedHeader.length) / 8) * 8;
    const mixed = new Uint8Array(mixedBase + mixedBytes.length - oldBase);
    mixed.set(mixedBytes.subarray(0, 4));
    new DataView(mixed.buffer).setUint32(4, mixedHeader.length, true);
    mixed.set(mixedHeader, 8);
    mixed.set(mixedBytes.subarray(oldBase), mixedBase);
    const mixedPage = decodeBlock(mixed);
    expect(mixedPage.cell(0)?.input).toEqual(expected[0]);
    expect(mixedPage.cell(1)?.input).toEqual(expected[1]);
    expect([...mixedPage.records()].map((cell) => cell.input)).toEqual([expected[1], expected[0]]);
  });
  it('rejects a position run before it exceeds its declared bounded count', () => {
    const cells: Record<string, import('@opensheetjs/core').CellRecord> = {};
    for (let column = 0; column < 32; column++)
      cells[key(0, column)] = {
        rowId: 0,
        columnId: column,
        input: { type: 'number', value: column },
      };
    const encoded = encodeBlock(cells, [], [], {});
    const view = new DataView(encoded.buffer, encoded.byteOffset, encoded.byteLength);
    const headerLength = view.getUint32(4, true);
    const metadata = JSON.parse(new TextDecoder().decode(encoded.subarray(8, 8 + headerLength)));
    const section = metadata.sections[0];
    expect(section.positions.mode).toBe(1);
    expect(section.positions.count).toBe(32);
    const base = Math.ceil((8 + headerLength) / 8) * 8;
    view.setUint16(base + section.positions.offset + 2, 2048, true);
    expect(() => decodeBlock(encoded)).toThrow('Invalid position run');
  });
  it('rejects fractional position counts before reading the position list', () => {
    const bytes = encodeBlock(
      { [key(0, 0)]: { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } } },
      [],
      [],
      {},
    );
    const length = new DataView(bytes.buffer, bytes.byteOffset).getUint32(4, true);
    const oldBase = Math.ceil((8 + length) / 8) * 8;
    const metadata = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + length)));
    expect(metadata.sections[0].positions.mode).toBe(0);
    metadata.sections[0].positions.count = 0.5;
    metadata.sections[0].positions.bytes = 1;
    const header = new TextEncoder().encode(JSON.stringify(metadata));
    const base = Math.ceil((8 + header.length) / 8) * 8;
    const damaged = new Uint8Array(base + bytes.length - oldBase);
    damaged.set(bytes.subarray(0, 4));
    new DataView(damaged.buffer).setUint32(4, header.length, true);
    damaged.set(header, 8);
    damaged.set(bytes.subarray(oldBase), base);
    expect(() => decodeBlock(damaged)).toThrow(
      expect.objectContaining({ code: 'CORRUPT_STORAGE' }),
    );
  });
  it('classifies invalid page JSON and UTF-8 as corrupt storage', () => {
    const bytes = encodeBlock(
      { [key(0, 0)]: { rowId: 0, columnId: 0, input: { type: 'string', value: 'hello' } } },
      [],
      [],
      {},
    );
    const headerLength = new DataView(bytes.buffer, bytes.byteOffset).getUint32(4, true);
    const meta = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + headerLength)));
    const section = meta.sections[0];
    const base = Math.ceil((8 + headerLength) / 8) * 8;
    expect(section.type).toBe(0);
    const textOffset = base + section.offset + 4 + 8;
    for (const [offset, byte] of [
      [8, 0x21],
      [8, 0xff],
      [textOffset, 0xff],
    ]) {
      const damaged = bytes.slice();
      damaged[offset] = byte;
      expect(() => decodeBlock(damaged)).toThrow(
        expect.objectContaining({ code: 'CORRUPT_STORAGE' }),
      );
    }
  });
  it('bounds dictionary expansion while decoding, before constructing all values', () => {
    const cells: Record<string, import('@opensheetjs/core').CellRecord> = {};
    for (let i = 0; i < 400; i++) {
      const row = Math.floor(i / 32),
        column = i % 32;
      cells[key(row, column)] = {
        rowId: row,
        columnId: column,
        input: { type: 'string', value: 'x'.repeat(32768) },
      };
    }
    const encoded = encodeBlock(cells, [], [], {});
    const decode = vi.spyOn(TextDecoder.prototype, 'decode');
    try {
      expect(() => decodeBlock(encoded)).toThrow('Page text exceeds 32 MiB working budget');
      expect(decode.mock.calls.length).toBeLessThan(400);
    } finally {
      decode.mockRestore();
    }
  });
  it('rejects oversized encoded strings before decoding their contents', () => {
    const encoded = encodeBlock(
      {
        [key(0, 0)]: {
          rowId: 0,
          columnId: 0,
          input: { type: 'string', value: 'x'.repeat(100000) },
        },
      },
      [],
      [],
      {},
    );
    const decode = vi.spyOn(TextDecoder.prototype, 'decode');
    try {
      expect(() => decodeBlock(encoded)).toThrow('Page string exceeds text limit');
      expect(decode).toHaveBeenCalledTimes(1);
    } finally {
      decode.mockRestore();
    }
  });
});

it('parses token batches across tiny UTF-8 chunks and releases a cancelled source', async () => {
  const book = await track();
  await book
    .getSheets()[0]
    .range('A1:B100')
    .setValues(Array.from({ length: 100 }, (_, i) => [i, `中😀\\"\n${i}`]));
  const file = await book.toJSON();
  const encoded = new TextEncoder().encode(JSON.stringify(file));
  for (const chunkSize of [1, 7, 127]) {
    let offset = 0;
    const source = new ReadableStream<Uint8Array>({
      pull(controller) {
        if (offset >= encoded.length) return controller.close();
        controller.enqueue(encoded.slice(offset, offset + chunkSize));
        offset += chunkSize;
      },
    });
    const blocks: unknown[] = [];
    let metadata: WorkbookFile | undefined;
    for await (const item of parseWorkbookJSON(source)) {
      if ('block' in item) blocks.push(item.block);
      else metadata = item.metadata;
    }
    expect(blocks).toEqual(file.sheets[0].blocks);
    expect(metadata?.sheets[0].blocks).toEqual([]);
  }
  let cancelled = false;
  const source = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(encoded);
    },
    cancel() {
      cancelled = true;
    },
  });
  for await (const item of parseWorkbookJSON(source)) {
    if ('block' in item) break;
  }
  expect(cancelled).toBe(true);
});

it('matches JSON.parse across seeded token batches and byte boundaries', async () => {
  const atom = fc.oneof(
    fc.integer({ min: -1_000_000, max: 1_000_000 }),
    fc.boolean(),
    fc.constant(null),
    fc
      .array(fc.constantFrom('A', '中', '😀', '\\', '"', '\n', '\r', '\t', '\ud800'), {
        maxLength: 12,
      })
      .map((parts) => parts.join('')),
  );
  await fc.assert(
    fc.asyncProperty(
      fc.array(atom, { minLength: 200, maxLength: 800 }),
      fc.oneof(fc.integer({ min: 1, max: 37 }), fc.constant(65536)),
      async (values, chunkSize) => {
        const text = JSON.stringify({ values });
        const bytes = new TextEncoder().encode(text);
        let position = 0;
        const source = new ReadableStream<Uint8Array>({
          pull(controller) {
            if (position >= bytes.length) return controller.close();
            controller.enqueue(bytes.slice(position, position + chunkSize));
            position += chunkSize;
          },
        });
        let result: unknown;
        for await (const item of parseWorkbookJSON(source)) {
          if ('metadata' in item) result = item.metadata;
          else throw new Error('Unexpected block');
        }
        expect(result).toEqual(JSON.parse(text));
      },
    ),
    { seed: 20261031, numRuns: 40 },
  );
}, 30_000);

it('checks cancellation between whitespace-only input chunks before any block is emitted', async () => {
  const controller = new AbortController();
  let cancelled = false,
    checkpoints = 0;
  const source = new ReadableStream<Uint8Array>({
    pull(stream) {
      stream.enqueue(new TextEncoder().encode(' '.repeat(2048)));
    },
    cancel() {
      cancelled = true;
    },
  });
  await expect(
    (async () => {
      for await (const _item of parseWorkbookJSON(source, {
        checkpoint: async () => {
          if (++checkpoints === 2) controller.abort();
          controller.signal.throwIfAborted();
        },
      })) {
        /* no block should be produced */
      }
    })(),
  ).rejects.toMatchObject({ name: 'AbortError' });
  expect(checkpoints).toBe(2);
  expect(cancelled).toBe(true);
});
