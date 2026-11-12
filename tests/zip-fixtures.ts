/** Two stored entries with identical content and a standard CRC, including proper offsets. */
export function twoEntryZip(names: [string, string]): Buffer {
  const body: Buffer[] = [];
  const directory: Buffer[] = [];
  let offset = 0;
  for (const name of names) {
    const filename = Buffer.from(name);
    const data = Buffer.from('abc');
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(0x352441c2, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(filename.length, 26);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(0x352441c2, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(filename.length, 28);
    central.writeUInt32LE(offset, 42);
    body.push(local, filename, data);
    directory.push(central, filename);
    offset += local.length + filename.length + data.length;
  }
  const index = Buffer.concat(directory);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(names.length, 8);
  end.writeUInt16LE(names.length, 10);
  end.writeUInt32LE(index.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...body, index, end]);
}

/** Append a legal ZIP comment that contains a misleading EOCD signature. */
export function withEndSignature(bytes: Buffer): Buffer {
  const comment = Buffer.alloc(32);
  comment.writeUInt32LE(0x06054b50, 4);
  const result = Buffer.concat([bytes, comment]);
  result.writeUInt16LE(comment.length, bytes.length - 2);
  return result;
}

/** Two distinct unflagged byte names that the downstream byte-wise lowercase folds together. */
export function byteCaseCollisionZip(): Buffer {
  const bytes = twoEntryZip(['x/aa', 'x/bb']);
  let directory = bytes.readUInt32LE(bytes.length - 6);
  for (const firstByte of [0xc4, 0xe4]) {
    const local = bytes.readUInt32LE(directory + 42);
    for (const offset of [directory + 46, local + 30]) {
      bytes[offset + 2] = firstByte;
      bytes[offset + 3] = 0xb0;
    }
    directory += 46 + bytes.readUInt16LE(directory + 28);
  }
  return bytes;
}

/** Convert the final stored ZIP entry to a size-less local descriptor. */
export function storedDescriptorZip(original: Buffer, signed = true): Buffer {
  const end = original.length - 22;
  const oldDirectory = original.readUInt32LE(end + 16);
  const count = original.readUInt16LE(end + 10);
  let cursor = oldDirectory;
  let selected = -1;
  let lastLocal = -1;
  for (let i = 0; i < count; i++) {
    const local = original.readUInt32LE(cursor + 42);
    if (local > lastLocal) {
      lastLocal = local;
      selected = cursor;
    }
    cursor +=
      46 +
      original.readUInt16LE(cursor + 28) +
      original.readUInt16LE(cursor + 30) +
      original.readUInt16LE(cursor + 32);
  }
  if (selected < 0 || original.readUInt16LE(selected + 10) !== 0)
    throw new Error('Fixture requires a stored ZIP entry');
  const size = original.readUInt32LE(selected + 20);
  const descriptor = Buffer.alloc(signed ? 16 : 12);
  if (signed) descriptor.writeUInt32LE(0x08074b50);
  const start = signed ? 4 : 0;
  descriptor.writeUInt32LE(original.readUInt32LE(selected + 16), start);
  descriptor.writeUInt32LE(size, start + 4);
  descriptor.writeUInt32LE(original.readUInt32LE(selected + 24), start + 8);
  const result = Buffer.concat([
    original.subarray(0, oldDirectory),
    descriptor,
    original.subarray(oldDirectory),
  ]);
  result.writeUInt16LE(result.readUInt16LE(lastLocal + 6) | 8, lastLocal + 6);
  result.writeUInt32LE(0, lastLocal + 14);
  result.writeUInt32LE(0, lastLocal + 18);
  result.writeUInt32LE(0, lastLocal + 22);
  const newSelected = selected + descriptor.length;
  result.writeUInt16LE(result.readUInt16LE(newSelected + 8) | 8, newSelected + 8);
  result.writeUInt32LE(oldDirectory + descriptor.length, result.length - 6);
  return result;
}
