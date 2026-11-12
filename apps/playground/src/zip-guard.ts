// ZIP record layout: https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
// This import profile accepts ordinary single-disk stored/deflated ZIPs, not ZIP64.
export async function validateZip(
  buffer: ArrayBuffer,
  required = false,
): Promise<Array<{ offset: number; size: number }>> {
  const bytes = new Uint8Array(buffer),
    view = new DataView(buffer);
  const fail = (message: string): never => {
    throw new Error(`Invalid ZIP: ${message}`);
  };
  const bounds = (offset: number, length: number, end = bytes.length) => {
    if (offset < 0 || offset + length > end) fail('truncated record');
  };
  if (bytes.length < 4 || view.getUint32(0, true) !== 0x04034b50) {
    if (required || (bytes[0] === 0x50 && bytes[1] === 0x4b)) fail('missing local header');
    return [];
  }
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (
      view.getUint32(i, true) === 0x06054b50 &&
      i + 22 + view.getUint16(i + 20, true) === bytes.length
    ) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) fail('missing end record');
  // SheetJS scans backwards for the last signature without validating its
  // position. A signature in the comment or fields would make it use a
  // different directory than the one checked here.
  for (let i = eocd + 4; i + 4 <= bytes.length; i++)
    if (view.getUint32(i, true) === 0x06054b50) fail('ambiguous end record');
  const count = view.getUint16(eocd + 10, true),
    directory = view.getUint32(eocd + 16, true);
  if (
    view.getUint16(eocd + 4, true) ||
    view.getUint16(eocd + 6, true) ||
    view.getUint16(eocd + 8, true) !== count
  )
    fail('multi-disk archives are unsupported');
  if (!count || count > 10000) fail('entry count exceeds import limits');
  if (directory + view.getUint32(eocd + 12, true) !== eocd) fail('directory bounds');
  const entries: {
    start: number;
    data: number;
    end: number;
    size: number;
    compressed: number;
    method: number;
    crc: number;
  }[] = [];
  const names = new Set<string>();
  const parserNames = new Set<string>();
  const storedSizeOffsets: Array<{ offset: number; size: number }> = [];
  let position = directory,
    expanded = 0;
  const extras = (start: number, length: number) => {
    const end = start + length;
    while (start < end) {
      bounds(start, 4, end);
      const kind = view.getUint16(start, true);
      if (kind === 1) fail('ZIP64 is unsupported');
      const size = view.getUint16(start + 2, true);
      bounds(start + 4, size, end);
      if (kind === 0x5455) {
        if (size < 1) fail('invalid extended timestamp');
        const flags = view.getUint8(start + 4);
        const required =
          1 + (flags & 1 ? 4 : 0) + (size > 5 ? (flags & 2 ? 4 : 0) + (flags & 4 ? 4 : 0) : 0);
        if (size < required) fail('invalid extended timestamp');
      }
      start += 4 + size;
    }
  };
  for (let i = 0; i < count; i++) {
    bounds(position, 46, eocd);
    if (view.getUint32(position, true) !== 0x02014b50) fail('directory signature');
    const flags = view.getUint16(position + 8, true),
      method = view.getUint16(position + 10, true);
    if (flags & ~0x080e || (method !== 0 && method !== 8))
      fail('unsupported encryption or compression');
    const crc = view.getUint32(position + 16, true),
      compressed = view.getUint32(position + 20, true),
      size = view.getUint32(position + 24, true);
    expanded += size;
    if (expanded > 200 * 1024 * 1024) fail('expanded file exceeds 200 MiB');
    const nameLength = view.getUint16(position + 28, true),
      extraLength = view.getUint16(position + 30, true),
      commentLength = view.getUint16(position + 32, true);
    bounds(position + 46, nameLength + extraLength + commentLength, eocd);
    extras(position + 46 + nameLength, extraLength);
    const nameBytes = bytes.subarray(position + 46, position + 46 + nameLength);
    let name: string;
    try {
      name = new TextDecoder('utf-8', { fatal: !!(flags & 0x0800) }).decode(nameBytes);
    } catch {
      return fail('invalid UTF-8 entry name');
    }
    const lookupName = name.toLowerCase();
    // The downstream ZIP reader currently turns each filename byte into a
    // character even when the UTF-8 flag is present, then lowercases it.
    let parserName = '';
    for (let offset = 0; offset < nameBytes.length; offset += 8192)
      parserName += String.fromCharCode(...nameBytes.subarray(offset, offset + 8192));
    parserName = parserName.toLowerCase();
    if (!name || name.includes('\0') || names.has(lookupName) || parserNames.has(parserName))
      fail('empty or duplicate entry name');
    const parts = name.split('/');
    if (
      name.startsWith('/') ||
      /^[a-z]:/i.test(name) ||
      /[\u0000-\u001f\u007f\\]/.test(name) ||
      parts.some(
        (part, index) => part === '.' || part === '..' || (!part && index !== parts.length - 1),
      )
    )
      fail('unsafe entry path');
    names.add(lookupName);
    parserNames.add(parserName);
    if (view.getUint16(position + 34, true)) fail('multi-disk entry');
    const start = view.getUint32(position + 42, true);
    bounds(start, 30, directory);
    if (
      view.getUint32(start, true) !== 0x04034b50 ||
      view.getUint16(start + 6, true) !== flags ||
      view.getUint16(start + 8, true) !== method
    )
      fail('local header disagrees with directory');
    const localNameLength = view.getUint16(start + 26, true),
      localExtraLength = view.getUint16(start + 28, true);
    bounds(start + 30, localNameLength + localExtraLength, directory);
    if (localNameLength !== nameLength || nameBytes.some((b, n) => bytes[start + 30 + n] !== b))
      fail('local entry name disagrees');
    extras(start + 30 + localNameLength, localExtraLength);
    const data = start + 30 + localNameLength + localExtraLength;
    bounds(data, compressed, directory);
    let end = data + compressed;
    if (flags & 8) {
      for (const [offset, expected] of [
        [14, crc],
        [18, compressed],
        [22, size],
      ]) {
        const declared = view.getUint32(start + offset, true);
        if (declared !== 0 && declared !== expected) fail('streaming local header disagrees');
      }
      bounds(end, 12, directory);
      const matches = (at: number) =>
        view.getUint32(at, true) === crc &&
        view.getUint32(at + 4, true) === compressed &&
        view.getUint32(at + 8, true) === size;
      // An unsigned descriptor's CRC can itself equal the optional signature.
      if (!matches(end) && view.getUint32(end, true) === 0x08074b50) end += 4;
      bounds(end, 12, directory);
      if (!matches(end)) fail('data descriptor disagrees');
      end += 12;
      // SheetJS uses the local compressed size to slice stored data, even
      // when bit 3 permits that size to be omitted until the descriptor.
      if (method === 0 && compressed && view.getUint32(start + 18, true) === 0)
        storedSizeOffsets.push({ offset: start + 18, size: compressed });
    } else if (
      view.getUint32(start + 14, true) !== crc ||
      view.getUint32(start + 18, true) !== compressed ||
      view.getUint32(start + 22, true) !== size
    )
      fail('local sizes or checksum disagree');
    if (method === 0 && compressed !== size) fail('stored size disagrees');
    entries.push({ start, data, end, size, compressed, method, crc });
    position += 46 + nameLength + extraLength + commentLength;
  }
  if (position !== eocd) fail('unexpected directory records');
  entries.sort((a, b) => a.start - b.start);
  let next = 0;
  for (const entry of entries) {
    if (entry.start !== next) fail('overlapping or unindexed local records');
    next = entry.end;
  }
  if (next !== directory) fail('unindexed data before directory');
  // Count real output, not attacker-controlled declared sizes. Discard each chunk.
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let n = i;
    for (let bit = 0; bit < 8; bit++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
    table[i] = n;
  }
  for (const entry of entries) {
    const payload = bytes.subarray(entry.data, entry.data + entry.compressed);
    let length = 0,
      crc = 0xffffffff;
    const consume = (chunk: Uint8Array) => {
      length += chunk.length;
      if (length > entry.size) fail('actual expanded size exceeds declared size');
      for (const byte of chunk) crc = table[(crc ^ byte) & 255] ^ (crc >>> 8);
    };
    if (entry.method === 0) consume(payload);
    else {
      let offset = 0;
      const input = new ReadableStream<Uint8Array<ArrayBuffer>>({
        pull(controller) {
          if (offset === payload.length) {
            controller.close();
            return;
          }
          // Bound compressed input supplied per native transform invocation.
          const end = Math.min(offset + 1024, payload.length);
          controller.enqueue(payload.subarray(offset, end));
          offset = end;
        },
      });
      const stream = input.pipeThrough(new DecompressionStream('deflate-raw'));
      const reader = stream.getReader();
      let finished = false;
      try {
        while (true) {
          const result = await reader.read();
          if (result.done) {
            finished = true;
            break;
          }
          consume(result.value);
        }
      } catch (error) {
        if (error instanceof Error && error.message.startsWith('Invalid ZIP:')) throw error;
        fail('invalid deflate stream');
      } finally {
        // The reader is private to this preflight. WebKit can hang while
        // releasing a completed decompression reader; no caller needs its lock.
        // The stream is owned here and cannot publish data. Do not let a
        // stalled native cancellation hide the validation error; the file
        // worker is terminated by its host after returning that error.
        if (!finished) void reader.cancel().catch(() => {});
      }
    }
    if (length !== entry.size || (crc ^ 0xffffffff) >>> 0 !== entry.crc)
      fail('expanded size or checksum disagrees');
  }
  return storedSizeOffsets;
}
