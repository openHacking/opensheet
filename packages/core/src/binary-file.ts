import { assert } from './types.js';
import { decodeFileMetadata, type WorkbookFile } from './file-codec.js';
export const crc32 = (bytes: Uint8Array) => {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
};
export async function gzipPage(bytes: Uint8Array): Promise<{ codec: number; bytes: Uint8Array }> {
  if (typeof CompressionStream === 'undefined') return { codec: 0, bytes };
  const encoded = new Uint8Array(
    await new Response(
      new Blob([bytes as BlobPart]).stream().pipeThrough(new CompressionStream('gzip')),
    ).arrayBuffer(),
  );
  return encoded.length <= bytes.length * 0.9 ? { codec: 1, bytes: encoded } : { codec: 0, bytes };
}
export async function unpackPage(
  bytes: Uint8Array,
  codec: number,
  decodedLength: number,
  limit = 32 * 1024 * 1024,
) {
  assert(
    Number.isInteger(decodedLength) && decodedLength > 0 && decodedLength <= limit,
    'LIMIT_EXCEEDED',
    'Decoded page exceeds budget',
  );
  if (codec === 0) {
    assert(bytes.length === decodedLength, 'INVALID_ARGUMENT', 'Page length differs');
    return bytes;
  }
  assert(
    codec === 1 && typeof DecompressionStream !== 'undefined',
    'INVALID_ARGUMENT',
    'Unsupported page codec',
  );
  const reader = new Blob([bytes as BlobPart])
    .stream()
    .pipeThrough(new DecompressionStream('gzip'))
    .getReader();
  const result = new Uint8Array(decodedLength);
  let offset = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      assert(
        offset + value.length <= decodedLength,
        'LIMIT_EXCEEDED',
        'Decompression exceeds declared size',
      );
      result.set(value, offset);
      offset += value.length;
    }
    assert(offset === decodedLength, 'INVALID_ARGUMENT', 'Decoded length differs');
  } finally {
    await reader.cancel().catch(() => {});
    reader.releaseLock();
  }
  return result;
}
type Entry = { sheet: number; row: number; column: number; offset: number; length: number };
export async function* writeBinary(
  metadata: WorkbookFile,
  pages: AsyncIterable<{ sheet: number; row: number; column: number; bytes: Uint8Array }>,
  signal?: AbortSignal,
): AsyncGenerator<Uint8Array> {
  const check = () => assert(!signal?.aborted, 'ABORTED', 'Export aborted');
  check();
  const text = new TextEncoder().encode(JSON.stringify(metadata));
  assert(text.length <= 32 * 1024 * 1024, 'LIMIT_EXCEEDED', 'Binary metadata exceeds 32 MiB');
  const header = new Uint8Array(12),
    hv = new DataView(header.buffer);
  header.set([79, 83, 66, 3]);
  hv.setUint32(4, text.length, true);
  hv.setUint32(8, crc32(text), true);
  yield header;
  yield text;
  let offset = header.length + text.length;
  const entries: Entry[] = [];
  for await (const page of pages) {
    check();
    const packed = await gzipPage(page.bytes);
    check();
    const frame = new Uint8Array(28),
      view = new DataView(frame.buffer);
    view.setUint32(0, page.sheet, true);
    view.setUint32(4, page.row, true);
    view.setUint32(8, page.column, true);
    view.setUint32(12, packed.codec, true);
    view.setUint32(16, packed.bytes.length, true);
    view.setUint32(20, page.bytes.length, true);
    view.setUint32(24, crc32(packed.bytes), true);
    entries.push({
      sheet: page.sheet,
      row: page.row,
      column: page.column,
      offset,
      length: 28 + packed.bytes.length,
    });
    yield frame;
    yield packed.bytes;
    offset += 28 + packed.bytes.length;
    assert(Number.isSafeInteger(offset), 'LIMIT_EXCEEDED', 'File offset exceeds safe integer');
  }
  check();
  const directory = new TextEncoder().encode(JSON.stringify(entries));
  assert(directory.length <= 32 * 1024 * 1024, 'LIMIT_EXCEEDED', 'Binary directory exceeds 32 MiB');
  const marker = new Uint8Array(8),
    mv = new DataView(marker.buffer);
  mv.setUint32(0, 0xffffffff, true);
  mv.setUint32(4, directory.length, true);
  yield marker;
  yield directory;
  const footer = new Uint8Array(24),
    fv = new DataView(footer.buffer);
  footer.set([79, 83, 66, 69]);
  fv.setBigUint64(4, BigInt(offset + 8), true);
  fv.setBigUint64(12, BigInt(directory.length), true);
  fv.setUint32(20, crc32(directory), true);
  check();
  yield footer;
}
class StreamReader {
  private reader: ReadableStreamDefaultReader<Uint8Array>;
  private chunk: Uint8Array = new Uint8Array(0);
  private at = 0;
  offset = 0;
  constructor(source: Blob | Uint8Array | ReadableStream<Uint8Array>) {
    this.reader = (
      source instanceof Blob
        ? source.stream()
        : source instanceof Uint8Array
          ? new Blob([source as BlobPart]).stream()
          : source
    ).getReader();
  }
  async read(length: number) {
    assert(
      Number.isInteger(length) && length >= 0 && length <= 32 * 1024 * 1024,
      'LIMIT_EXCEEDED',
      'File section exceeds 32 MiB',
    );
    const result = new Uint8Array(length);
    let at = 0;
    while (at < length) {
      if (this.at === this.chunk.length) {
        const item = await this.reader.read();
        assert(!item.done, 'INVALID_ARGUMENT', 'Truncated binary file');
        this.chunk = item.value;
        this.at = 0;
      }
      const count = Math.min(length - at, this.chunk.length - this.at);
      result.set(this.chunk.subarray(this.at, this.at + count), at);
      at += count;
      this.at += count;
      this.offset += count;
    }
    return result;
  }
  async end() {
    assert(
      this.at === this.chunk.length && (await this.reader.read()).done,
      'INVALID_ARGUMENT',
      'Trailing binary data',
    );
  }
  async close() {
    await this.reader.cancel().catch(() => {});
    this.reader.releaseLock();
  }
}
export async function* readBinary(
  source: Blob | Uint8Array | ReadableStream<Uint8Array>,
): AsyncGenerator<
  { metadata: WorkbookFile } | { sheet: number; row: number; column: number; bytes: Uint8Array }
> {
  const reader = new StreamReader(source),
    decoder = new TextDecoder('utf-8', { fatal: true });
  const entries: Entry[] = [];
  try {
    const header = await reader.read(12),
      hv = new DataView(header.buffer);
    assert(
      header[0] === 79 && header[1] === 83 && header[2] === 66 && header[3] === 3,
      'INVALID_ARGUMENT',
      'Unsupported binary format',
    );
    const text = await reader.read(hv.getUint32(4, true));
    assert(crc32(text) === hv.getUint32(8, true), 'INVALID_ARGUMENT', 'Header checksum differs');
    yield { metadata: JSON.parse(decoder.decode(text)) };
    while (true) {
      const offset = reader.offset,
        first = await reader.read(4),
        sheet = new DataView(first.buffer).getUint32(0, true);
      if (sheet === 0xffffffff) {
        const size = new DataView((await reader.read(4)).buffer).getUint32(0, true),
          directoryOffset = reader.offset,
          directory = await reader.read(size),
          footer = await reader.read(24),
          fv = new DataView(footer.buffer);
        assert(
          footer[0] === 79 &&
            footer[1] === 83 &&
            footer[2] === 66 &&
            footer[3] === 69 &&
            fv.getBigUint64(4, true) === BigInt(directoryOffset) &&
            fv.getBigUint64(12, true) === BigInt(size) &&
            fv.getUint32(20, true) === crc32(directory),
          'INVALID_ARGUMENT',
          'Invalid file directory',
        );
        assert(
          JSON.stringify(JSON.parse(decoder.decode(directory))) === JSON.stringify(entries),
          'INVALID_ARGUMENT',
          'Directory differs from pages',
        );
        await reader.end();
        break;
      }
      const rest = await reader.read(24),
        view = new DataView(rest.buffer),
        row = view.getUint32(0, true),
        column = view.getUint32(4, true),
        codec = view.getUint32(8, true),
        length = view.getUint32(12, true),
        decodedLength = view.getUint32(16, true),
        checksum = view.getUint32(20, true);
      const bytes = await reader.read(length);
      assert(crc32(bytes) === checksum, 'INVALID_ARGUMENT', 'Page checksum differs');
      entries.push({ sheet, row, column, offset, length: 28 + length });
      yield { sheet, row, column, bytes: await unpackPage(bytes, codec, decodedLength) };
    }
  } finally {
    await reader.close();
  }
}
/** Independent page access for browser and Node Blob/byte readers. */
export class BinaryReader {
  private constructor(
    private source: Blob,
    readonly metadata: WorkbookFile,
    readonly directory: Entry[],
  ) {}
  static async open(input: Blob | Uint8Array) {
    const source = input instanceof Blob ? input : new Blob([input as BlobPart]);
    assert(source.size >= 36, 'INVALID_ARGUMENT', 'Truncated binary file');
    const footer = new Uint8Array(await source.slice(-24).arrayBuffer()),
      fv = new DataView(footer.buffer),
      offset = Number(fv.getBigUint64(4, true)),
      length = Number(fv.getBigUint64(12, true));
    assert(
      footer[0] === 79 &&
        footer[1] === 83 &&
        footer[2] === 66 &&
        footer[3] === 69 &&
        Number.isSafeInteger(offset) &&
        Number.isSafeInteger(length) &&
        length <= 32 * 1024 * 1024 &&
        offset + length === source.size - 24,
      'INVALID_ARGUMENT',
      'Invalid directory bounds',
    );
    const directory = new Uint8Array(await source.slice(offset, offset + length).arrayBuffer());
    assert(
      crc32(directory) === fv.getUint32(20, true),
      'INVALID_ARGUMENT',
      'Directory checksum differs',
    );
    const header = new Uint8Array(await source.slice(0, 12).arrayBuffer()),
      hv = new DataView(header.buffer);
    assert(
      header[0] === 79 && header[1] === 83 && header[2] === 66 && header[3] === 3,
      'INVALID_ARGUMENT',
      'Unsupported binary format',
    );
    const size = hv.getUint32(4, true);
    assert(
      size <= 32 * 1024 * 1024 && 12 + size <= offset - 8,
      'INVALID_ARGUMENT',
      'Invalid metadata bounds',
    );
    const text = new Uint8Array(await source.slice(12, 12 + size).arrayBuffer());
    assert(crc32(text) === hv.getUint32(8, true), 'INVALID_ARGUMENT', 'Header checksum differs');
    const metadata = JSON.parse(
      new TextDecoder('utf-8', { fatal: true }).decode(text),
    ) as WorkbookFile;
    decodeFileMetadata(metadata);
    const seen = new Set<string>();
    const entries = JSON.parse(
      new TextDecoder('utf-8', { fatal: true }).decode(directory),
    ) as Entry[];
    assert(Array.isArray(entries), 'INVALID_ARGUMENT', 'Invalid directory');
    let end = 12 + size;
    for (const entry of entries) {
      assert(
        Number.isInteger(entry.sheet) &&
          entry.sheet >= 0 &&
          entry.sheet < metadata.sheets.length &&
          Number.isInteger(entry.row) &&
          entry.row >= 0 &&
          Number.isInteger(entry.column) &&
          entry.column >= 0 &&
          entry.offset === end &&
          Number.isInteger(entry.length) &&
          entry.length >= 28 &&
          entry.offset + entry.length <= offset - 8,
        'INVALID_ARGUMENT',
        'Invalid directory entry',
      );
      const address = `${entry.sheet}/${entry.row}/${entry.column}`;
      assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate directory page');
      seen.add(address);
      end += entry.length;
    }
    assert(end === offset - 8, 'INVALID_ARGUMENT', 'Invalid page coverage');
    return new BinaryReader(source, metadata, entries);
  }
  async page(index: number) {
    const entry = this.directory[index];
    assert(entry, 'INVALID_ARGUMENT', 'Unknown page');
    const bytes = new Uint8Array(
        await this.source.slice(entry.offset, entry.offset + entry.length).arrayBuffer(),
      ),
      view = new DataView(bytes.buffer);
    assert(
      view.getUint32(0, true) === entry.sheet &&
        view.getUint32(4, true) === entry.row &&
        view.getUint32(8, true) === entry.column &&
        view.getUint32(16, true) === bytes.length - 28 &&
        crc32(bytes.subarray(28)) === view.getUint32(24, true),
      'INVALID_ARGUMENT',
      'Invalid page frame',
    );
    return unpackPage(bytes.subarray(28), view.getUint32(12, true), view.getUint32(20, true));
  }
}
