import { assert, OpenSheetError, type WorkbookSnapshot } from './types.js';
import { yieldToTask } from './task-yield.js';
import { decodeFileMetadata, type WorkbookFile } from './file-codec.js';
function parseBinaryJSON(bytes: Uint8Array, part: 'metadata' | 'directory'): unknown {
  try {
    return JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(bytes));
  } catch {
    throw new OpenSheetError('INVALID_ARGUMENT', `Invalid binary ${part} JSON or UTF-8`);
  }
}
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
function validatePageEncoding(
  codec: number,
  encodedLength: number,
  decodedLength: number,
  limit: number,
) {
  assert(
    Number.isInteger(decodedLength) && decodedLength > 0 && decodedLength <= limit,
    'LIMIT_EXCEEDED',
    'Decoded page exceeds budget',
  );
  if (codec === 0) {
    assert(encodedLength === decodedLength, 'INVALID_ARGUMENT', 'Page length differs');
  } else {
    assert(
      codec === 1 && typeof DecompressionStream !== 'undefined',
      'INVALID_ARGUMENT',
      'Unsupported page codec',
    );
  }
}
export async function unpackPage(
  bytes: Uint8Array,
  codec: number,
  decodedLength: number,
  limit = 32 * 1024 * 1024,
  signal?: AbortSignal,
) {
  signal?.throwIfAborted();
  validatePageEncoding(codec, bytes.length, decodedLength, limit);
  if (codec === 0) {
    return bytes;
  }
  const reader = new Blob([bytes as BlobPart])
    .stream()
    .pipeThrough(new DecompressionStream('gzip'))
    .getReader();
  const result = new Uint8Array(decodedLength);
  let offset = 0;
  let finished = false;
  const abort = () => {
    void reader.cancel().catch(() => {});
  };
  signal?.addEventListener('abort', abort, { once: true });
  try {
    while (true) {
      signal?.throwIfAborted();
      let chunk: ReadableStreamReadResult<Uint8Array>;
      try {
        chunk = await reader.read();
      } catch {
        signal?.throwIfAborted();
        throw new OpenSheetError('CORRUPT_STORAGE', 'Invalid compressed page');
      }
      signal?.throwIfAborted();
      const { done, value } = chunk;
      if (done) {
        finished = true;
        break;
      }
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
    signal?.removeEventListener('abort', abort);
    // The stream and reader are private to this function. WebKit can hang in
    // releaseLock() even after decompression has finished; there is no caller
    // that needs the lock returned. Only cancel an unfinished stream.
    if (!finished) {
      const cancellation = reader.cancel().catch(() => {});
      if (!signal?.aborted) await cancellation;
    }
  }
  return result;
}
function assertBinaryHeader(header: Uint8Array) {
  assert(
    header[0] === 79 && header[1] === 83 && header[2] === 66,
    'INVALID_ARGUMENT',
    'Invalid binary workbook signature',
  );
  assert(
    header[3] === 3,
    'UNSUPPORTED_VERSION',
    `Binary workbook version ${header[3]} is unsupported; open it with a compatible version and export a schema 3 backup`,
  );
}
type Entry = { sheet: number; row: number; column: number; offset: number; length: number };
function hasDirectoryFields(value: unknown): value is Entry {
  return (
    value !== null &&
    typeof value === 'object' &&
    !Array.isArray(value) &&
    Object.keys(value).length === 5 &&
    Object.hasOwn(value, 'sheet') &&
    Object.hasOwn(value, 'row') &&
    Object.hasOwn(value, 'column') &&
    Object.hasOwn(value, 'offset') &&
    Object.hasOwn(value, 'length')
  );
}
const DIRECTORY_ENTRY_OVERHEAD = '{"sheet":,"row":,"column":,"offset":,"length":}'.length;
/** Track the exact ASCII JSON directory size before retaining another entry. */
export class BinaryDirectoryBudget {
  private bytes = 2; // Array brackets.
  private count = 0;
  constructor(private limit = 32 * 1024 * 1024) {}
  add(entry: Entry) {
    const next =
      this.bytes +
      (this.count ? 1 : 0) +
      DIRECTORY_ENTRY_OVERHEAD +
      String(entry.sheet).length +
      String(entry.row).length +
      String(entry.column).length +
      String(entry.offset).length +
      String(entry.length).length;
    assert(next <= this.limit, 'LIMIT_EXCEEDED', 'Binary directory exceeds 32 MiB');
    this.bytes = next;
    this.count++;
  }
}
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
  const directoryBudget = new BinaryDirectoryBudget();
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
    const entry = {
      sheet: page.sheet,
      row: page.row,
      column: page.column,
      offset,
      length: 28 + packed.bytes.length,
    };
    directoryBudget.add(entry);
    entries.push(entry);
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
  private chunksSinceCheck = 0;
  private lastYield = performance.now();
  offset = 0;
  private abort = () => {
    void this.reader.cancel().catch(() => {});
  };
  constructor(
    source: Blob | Uint8Array | ReadableStream<Uint8Array>,
    private signal?: AbortSignal,
  ) {
    signal?.throwIfAborted();
    this.reader = (
      source instanceof Blob
        ? source.stream()
        : source instanceof Uint8Array
          ? new Blob([source as BlobPart]).stream()
          : source
    ).getReader();
    signal?.addEventListener('abort', this.abort, { once: true });
  }
  private async nextChunk(): Promise<Uint8Array | undefined> {
    while (true) {
      const item = await this.reader.read();
      this.signal?.throwIfAborted();
      if (item.done) return undefined;
      assert(
        ArrayBuffer.isView(item.value) &&
          Object.prototype.toString.call(item.value) === '[object Uint8Array]',
        'INVALID_ARGUMENT',
        'Binary stream must contain Uint8Array chunks',
      );
      // Both empty and tiny nonempty chunks must allow timer-driven cancellation.
      if (++this.chunksSinceCheck >= 64) {
        this.chunksSinceCheck = 0;
        if (performance.now() - this.lastYield >= 8) {
          await yieldToTask();
          this.lastYield = performance.now();
          this.signal?.throwIfAborted();
        }
      }
      if (item.value.length) return item.value;
    }
  }
  async read(length: number) {
    this.signal?.throwIfAborted();
    assert(
      Number.isInteger(length) && length >= 0 && length <= 32 * 1024 * 1024,
      'LIMIT_EXCEEDED',
      'File section exceeds 32 MiB',
    );
    const result = new Uint8Array(length);
    let at = 0;
    while (at < length) {
      if (this.at === this.chunk.length) {
        const chunk = await this.nextChunk();
        assert(chunk, 'INVALID_ARGUMENT', 'Truncated binary file');
        this.chunk = chunk;
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
    this.signal?.throwIfAborted();
    const done = this.at === this.chunk.length && !(await this.nextChunk());
    this.signal?.throwIfAborted();
    assert(done, 'INVALID_ARGUMENT', 'Trailing binary data');
  }
  async close() {
    this.signal?.removeEventListener('abort', this.abort);
    const cancellation = this.reader.cancel().catch(() => {});
    if (!this.signal?.aborted) await cancellation;
    this.reader.releaseLock();
  }
}
export async function* readBinary(
  source: Blob | Uint8Array | ReadableStream<Uint8Array>,
  options: { signal?: AbortSignal } = {},
): AsyncGenerator<
  | { metadata: WorkbookFile; snapshot: WorkbookSnapshot }
  | { sheet: number; row: number; column: number; bytes: Uint8Array }
> {
  const reader = new StreamReader(source, options.signal);
  const entries: Entry[] = [];
  const seen = new Set<string>();
  const directoryBudget = new BinaryDirectoryBudget();
  try {
    const header = await reader.read(12),
      hv = new DataView(header.buffer);
    assertBinaryHeader(header);
    const text = await reader.read(hv.getUint32(4, true));
    assert(crc32(text) === hv.getUint32(8, true), 'INVALID_ARGUMENT', 'Header checksum differs');
    const metadata = parseBinaryJSON(text, 'metadata') as WorkbookFile;
    const snapshot = decodeFileMetadata(metadata);
    assert(
      metadata.sheets.every((sheet) => sheet.blocks.length === 0),
      'INVALID_ARGUMENT',
      'Binary metadata contains inline blocks',
    );
    yield { metadata, snapshot };
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
        const listed = parseBinaryJSON(directory, 'directory');
        assert(
          Array.isArray(listed) && listed.length === entries.length,
          'INVALID_ARGUMENT',
          'Directory differs from pages',
        );
        for (let index = 0; index < entries.length; index++) {
          const item = listed[index],
            expected = entries[index];
          assert(
            hasDirectoryFields(item) &&
              item.sheet === expected.sheet &&
              item.row === expected.row &&
              item.column === expected.column &&
              item.offset === expected.offset &&
              item.length === expected.length,
            'INVALID_ARGUMENT',
            'Directory differs from pages',
          );
        }
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
      assert(sheet < metadata.sheets.length, 'INVALID_ARGUMENT', 'Unknown binary page sheet');
      const address = `${sheet}/${row}/${column}`;
      assert(!seen.has(address), 'INVALID_ARGUMENT', 'Duplicate binary page');
      assert(Number.isSafeInteger(offset), 'LIMIT_EXCEEDED', 'File offset exceeds safe integer');
      const entry = { sheet, row, column, offset, length: 28 + length };
      directoryBudget.add(entry);
      seen.add(address);
      validatePageEncoding(codec, length, decodedLength, 32 * 1024 * 1024);
      const bytes = await reader.read(length);
      assert(crc32(bytes) === checksum, 'INVALID_ARGUMENT', 'Page checksum differs');
      entries.push(entry);
      yield {
        sheet,
        row,
        column,
        bytes: await unpackPage(bytes, codec, decodedLength, undefined, options.signal),
      };
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
    const header = new Uint8Array(await source.slice(0, 12).arrayBuffer()),
      hv = new DataView(header.buffer);
    assertBinaryHeader(header);
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
        offset >= 8 &&
        length <= 32 * 1024 * 1024 &&
        offset + length === source.size - 24,
      'INVALID_ARGUMENT',
      'Invalid directory bounds',
    );
    const marker = new DataView(await source.slice(offset - 8, offset).arrayBuffer());
    assert(
      marker.getUint32(0, true) === 0xffffffff && marker.getUint32(4, true) === length,
      'INVALID_ARGUMENT',
      'Invalid directory marker',
    );
    const directory = new Uint8Array(await source.slice(offset, offset + length).arrayBuffer());
    assert(
      crc32(directory) === fv.getUint32(20, true),
      'INVALID_ARGUMENT',
      'Directory checksum differs',
    );

    const size = hv.getUint32(4, true);
    assert(
      size <= 32 * 1024 * 1024 && 12 + size <= offset - 8,
      'INVALID_ARGUMENT',
      'Invalid metadata bounds',
    );
    const text = new Uint8Array(await source.slice(12, 12 + size).arrayBuffer());
    assert(crc32(text) === hv.getUint32(8, true), 'INVALID_ARGUMENT', 'Header checksum differs');
    const metadata = parseBinaryJSON(text, 'metadata') as WorkbookFile;
    decodeFileMetadata(metadata);
    assert(
      metadata.sheets.every((sheet) => sheet.blocks.length === 0),
      'INVALID_ARGUMENT',
      'Binary metadata contains inline blocks',
    );
    const seen = new Set<string>();
    const entries = parseBinaryJSON(directory, 'directory') as Entry[];
    assert(Array.isArray(entries), 'INVALID_ARGUMENT', 'Invalid directory');
    let end = 12 + size;
    for (const entry of entries) {
      assert(
        hasDirectoryFields(entry) &&
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
      assert(
        entry.length <= 32 * 1024 * 1024 + 28,
        'LIMIT_EXCEEDED',
        'Encoded page exceeds 32 MiB',
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
    assert(Number.isInteger(index) && index >= 0, 'INVALID_ARGUMENT', 'Unknown page');
    const storedEntry = this.directory[index];
    assert(storedEntry, 'INVALID_ARGUMENT', 'Unknown page');
    assert(hasDirectoryFields(storedEntry), 'INVALID_ARGUMENT', 'Invalid directory entry');
    const entry = { ...storedEntry };
    assert(
      Number.isSafeInteger(entry.offset) &&
        entry.offset >= 0 &&
        Number.isSafeInteger(entry.length) &&
        entry.length >= 28 &&
        entry.offset + entry.length <= this.source.size,
      'INVALID_ARGUMENT',
      'Invalid directory entry',
    );
    assert(entry.length <= 32 * 1024 * 1024 + 28, 'LIMIT_EXCEEDED', 'Encoded page exceeds 32 MiB');
    // Avoid an extra asynchronous Blob read for small pages, while rejecting
    // invalid large frames before materializing their payload.
    const headerOnly = entry.length > 64 * 1024;
    const bytes = new Uint8Array(
        await this.source
          .slice(entry.offset, entry.offset + (headerOnly ? 28 : entry.length))
          .arrayBuffer(),
      ),
      view = new DataView(bytes.buffer);
    assert(
      view.getUint32(0, true) === entry.sheet &&
        view.getUint32(4, true) === entry.row &&
        view.getUint32(8, true) === entry.column &&
        view.getUint32(16, true) === entry.length - 28,
      'INVALID_ARGUMENT',
      'Invalid page frame',
    );
    const codec = view.getUint32(12, true),
      decodedLength = view.getUint32(20, true);
    validatePageEncoding(codec, entry.length - 28, decodedLength, 32 * 1024 * 1024);
    const payload = headerOnly
      ? new Uint8Array(
          await this.source.slice(entry.offset + 28, entry.offset + entry.length).arrayBuffer(),
        )
      : bytes.subarray(28);
    assert(crc32(payload) === view.getUint32(24, true), 'INVALID_ARGUMENT', 'Invalid page frame');
    return unpackPage(payload, codec, decodedLength);
  }
}
