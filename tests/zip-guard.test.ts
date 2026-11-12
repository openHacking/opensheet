import { readFileSync } from 'node:fs';
import { deflateRawSync } from 'node:zlib';
import fc from 'fast-check';
import { expect, it, vi } from 'vitest';
import * as XLSX from 'xlsx';
import { validateZip } from '../apps/playground/src/zip-guard.js';
import {
  byteCaseCollisionZip,
  storedDescriptorZip,
  twoEntryZip,
  withEndSignature,
} from './zip-fixtures.js';

function archive(compression = true): ArrayBuffer {
  return XLSX.write(
    XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([['x'.repeat(32000), 7]]), 'Data'),
    {
      type: 'array',
      bookType: 'xlsx',
      compression,
    },
  );
}
function central(buffer: ArrayBuffer) {
  const view = new DataView(buffer);
  return view.getUint32(buffer.byteLength - 6, true);
}
function streamedEntry(
  signed: boolean,
  source = Buffer.from('abc'),
  crc = 0x352441c2,
  entryName = 'entry',
  extra = Buffer.alloc(0),
): ArrayBuffer {
  const data = deflateRawSync(source),
    name = Buffer.from(entryName);
  const local = Buffer.alloc(30),
    directory = Buffer.alloc(46),
    end = Buffer.alloc(22);
  const descriptor = Buffer.alloc(signed ? 16 : 12);
  local.writeUInt32LE(0x04034b50);
  local.writeUInt16LE(20, 4);
  local.writeUInt16LE(8, 6);
  local.writeUInt16LE(8, 8);
  local.writeUInt16LE(name.length, 26);
  local.writeUInt16LE(extra.length, 28);
  if (signed) descriptor.writeUInt32LE(0x08074b50);
  descriptor.writeUInt32LE(crc, signed ? 4 : 0);
  descriptor.writeUInt32LE(data.length, signed ? 8 : 4);
  descriptor.writeUInt32LE(source.length, signed ? 12 : 8);
  directory.writeUInt32LE(0x02014b50);
  directory.writeUInt16LE(20, 6);
  directory.writeUInt16LE(8, 8);
  directory.writeUInt16LE(8, 10);
  directory.writeUInt32LE(crc, 16);
  directory.writeUInt32LE(data.length, 20);
  directory.writeUInt32LE(source.length, 24);
  directory.writeUInt16LE(name.length, 28);
  directory.writeUInt16LE(extra.length, 30);
  const offset = local.length + name.length + extra.length + data.length + descriptor.length;
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(1, 8);
  end.writeUInt16LE(1, 10);
  end.writeUInt32LE(directory.length + name.length + extra.length, 12);
  end.writeUInt32LE(offset, 16);
  const bytes = Buffer.concat([local, name, extra, data, descriptor, directory, name, extra, end]);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}
it('does not release or cancel a completed private ZIP decompression reader', async () => {
  const Native = globalThis.DecompressionStream;
  let released = false,
    cancelled = false;
  class TrackedDecompression extends Native {
    constructor(format: CompressionFormat) {
      super(format);
      const readable = this.readable;
      const getReader = readable.getReader.bind(readable);
      readable.getReader = (() => {
        const reader = getReader();
        reader.releaseLock = () => {
          released = true;
          throw new Error('releaseLock must not be used');
        };
        const cancel = reader.cancel.bind(reader);
        reader.cancel = (...args) => {
          cancelled = true;
          return cancel(...args);
        };
        return reader;
      }) as typeof readable.getReader;
    }
  }
  vi.stubGlobal('DecompressionStream', TrackedDecompression);
  try {
    await validateZip(streamedEntry(true), true);
    expect(released).toBe(false);
    expect(cancelled).toBe(false);
  } finally {
    vi.unstubAllGlobals();
  }
});
it('returns a ZIP validation error when private stream cancellation never settles', async () => {
  const Native = globalThis.DecompressionStream;
  let cancelled = false;
  class StalledCancelDecompression extends Native {
    constructor(format: CompressionFormat) {
      super(format);
      const readable = this.readable;
      const getReader = readable.getReader.bind(readable);
      readable.getReader = (() => {
        const reader = getReader();
        reader.cancel = () => {
          cancelled = true;
          return new Promise<void>(() => {});
        };
        return reader;
      }) as typeof readable.getReader;
    }
  }
  const bytes = archive(),
    view = new DataView(bytes),
    directory = central(bytes);
  view.setUint32(directory + 24, 1, true);
  view.setUint32(view.getUint32(directory + 42, true) + 22, 1, true);
  vi.stubGlobal('DecompressionStream', StalledCancelDecompression);
  try {
    await expect(validateZip(bytes)).rejects.toThrow('actual expanded size');
    expect(cancelled).toBe(true);
  } finally {
    vi.unstubAllGlobals();
  }
}, 1000);
it('validates extended timestamp payloads before ZIP parsing', async () => {
  const extra = (payload: number[]) => {
    const header = Buffer.alloc(4);
    header.writeUInt16LE(0x5455);
    header.writeUInt16LE(payload.length, 2);
    return Buffer.concat([header, Buffer.from(payload)]);
  };
  for (const payload of [[], [1], [7, 0, 0, 0, 0, 0]])
    await expect(
      validateZip(
        streamedEntry(true, Buffer.from('abc'), 0x352441c2, 'entry', extra(payload)),
        true,
      ),
    ).rejects.toThrow('invalid extended timestamp');
  for (const payload of [[0], [1, 0, 0, 0, 0], [7, 0, 0, 0, 0]])
    await validateZip(
      streamedEntry(true, Buffer.from('abc'), 0x352441c2, 'entry', extra(payload)),
      true,
    );
});
it('rejects ambiguous and traversing entry paths', async () => {
  for (const name of [
    '../entry',
    'C:/entry',
    'c:entry',
    'xl/../entry',
    '/xl/entry',
    'xl\\entry',
    'xl//entry',
    'xl/./entry',
    'xl/\u0001entry',
  ])
    await expect(
      validateZip(streamedEntry(true, Buffer.from('abc'), 0x352441c2, name), true),
    ).rejects.toThrow('unsafe entry path');
  await validateZip(
    streamedEntry(true, Buffer.from('abc'), 0x352441c2, 'xl/worksheets/sheet1.xml'),
    true,
  );
});
it('rejects names that collide under the workbook parser case rules', async () => {
  const arrayBuffer = (bytes: Buffer) => Uint8Array.from(bytes).buffer;
  await expect(
    validateZip(arrayBuffer(twoEntryZip(['xl/sheet.xml', 'XL/SHEET.XML'])), true),
  ).rejects.toThrow('duplicate entry name');
  await validateZip(arrayBuffer(twoEntryZip(['xl/sheet1.xml', 'xl/sheet2.xml'])), true);
  await expect(validateZip(arrayBuffer(byteCaseCollisionZip()), true)).rejects.toThrow(
    'duplicate entry name',
  );
});
it('rejects end signatures that would redirect the downstream parser', async () => {
  const bytes = Buffer.from(archive());
  await expect(validateZip(Uint8Array.from(withEndSignature(bytes)).buffer, true)).rejects.toThrow(
    'ambiguous end record',
  );
  const comment = Buffer.from('ordinary comment');
  const accepted = Buffer.concat([bytes, comment]);
  accepted.writeUInt16LE(comment.length, bytes.length - 2);
  await validateZip(Uint8Array.from(accepted).buffer, true);
});
it('rejects invalid names declared as UTF-8', async () => {
  const bytes = streamedEntry(true);
  const view = new DataView(bytes);
  const directory = central(bytes);
  view.setUint16(6, 0x0808, true);
  view.setUint16(directory + 8, 0x0808, true);
  view.setUint8(30, 0xff);
  view.setUint8(directory + 46, 0xff);
  await expect(validateZip(bytes, true)).rejects.toThrow('Invalid ZIP: invalid UTF-8 entry name');
});
it('supports signed and unsigned streaming descriptors and rejects contradictory local sizes', async () => {
  for (const signed of [false, true]) {
    const bytes = streamedEntry(signed);
    await validateZip(bytes);
    new DataView(bytes).setUint32(22, 999, true);
    await expect(validateZip(bytes)).rejects.toThrow('streaming local header disagrees');
  }
});
it.each([true, false])('returns a parser fixup for a stored entry (signed=%s)', async (signed) => {
  const bytes = storedDescriptorZip(Buffer.from(archive(false)), signed);
  const buffer = Uint8Array.from(bytes).buffer;
  const sizes = await validateZip(buffer, true);
  expect(sizes).toHaveLength(1);
  expect(new DataView(buffer).getUint32(sizes[0].offset, true)).toBe(0);
  expect(() => XLSX.CFB.read(bytes, { type: 'buffer' })).toThrow('Bad compressed size');
  new DataView(buffer).setUint32(sizes[0].offset, sizes[0].size, true);
  const workbook = XLSX.read(buffer, { type: 'array' });
  expect(workbook.Sheets.Data.A1.v).toBe('x'.repeat(32000));
  const invalid = storedDescriptorZip(Buffer.from(archive(false)), signed);
  invalid[31] ^= 1;
  await expect(validateZip(Uint8Array.from(invalid).buffer, true)).rejects.toThrow('Invalid ZIP');
});
it('rejects name, disk, directory and descriptor inconsistencies', async () => {
  for (const mutate of [
    (v: DataView) => v.setUint8(30, 0),
    (v: DataView) => v.setUint16(v.byteLength - 18, 1, true),
    (v: DataView) => v.setUint32(v.byteLength - 10, 0, true),
    (v: DataView) => v.setUint32(central(v.buffer as ArrayBuffer) - 4, 99, true),
  ]) {
    const bytes = streamedEntry(true);
    mutate(new DataView(bytes));
    await expect(validateZip(bytes)).rejects.toThrow('Invalid ZIP');
  }
});
it('accepts stored, deflated and independently authored XLSX containers', async () => {
  await validateZip(archive(false), true);
  await validateZip(archive(), true);
  const bytes = readFileSync('benchmarks/interop-reader-2026-10-23-type-probe.authored.xlsx');
  await validateZip(
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    true,
  );
});
it('checks ZIP contents even without the XLSX extension and rejects truncated containers', async () => {
  const bytes = archive();
  await expect(validateZip(bytes.slice(0, -1))).rejects.toThrow('Invalid ZIP');
  await expect(validateZip(new ArrayBuffer(0), true)).rejects.toThrow('Invalid ZIP');
  await validateZip(new TextEncoder().encode('a,b\n1,2').buffer);
});
it('rejects forged expansion lengths before the workbook parser runs', async () => {
  const bytes = archive(),
    view = new DataView(bytes),
    directory = central(bytes);
  const local = view.getUint32(directory + 42, true);
  view.setUint32(directory + 24, 1, true);
  view.setUint32(local + 22, 1, true);
  await expect(validateZip(bytes)).rejects.toThrow('actual expanded size');
});
it('rejects metadata bombs, encryption, local/directory disagreement and CRC corruption', async () => {
  for (const mutation of [
    (v: DataView, d: number) => v.setUint32(d + 24, 201 * 1024 * 1024, true),
    (v: DataView, d: number) => v.setUint16(d + 8, 1, true),
    (v: DataView, d: number) => v.setUint32(d + 20, 0, true),
    (v: DataView, d: number) => {
      v.setUint32(d + 16, 0, true);
      v.setUint32(v.getUint32(d + 42, true) + 14, 0, true);
    },
  ]) {
    const bytes = archive();
    mutation(new DataView(bytes), central(bytes));
    await expect(validateZip(bytes)).rejects.toThrow('Invalid ZIP');
  }
});

it('distinguishes a descriptor signature from the same bytes used as a CRC', async () => {
  for (const signed of [false, true])
    await validateZip(streamedEntry(signed, Buffer.from('ac0a7ad5', 'hex'), 0x08074b50));
});

it('reports malformed deflate data through a stable ZIP error', async () => {
  const bytes = streamedEntry(true);
  new Uint8Array(bytes)[35] = 7; // BFINAL=1, reserved BTYPE=3.
  await expect(validateZip(bytes, true)).rejects.toThrow('Invalid ZIP: invalid deflate stream');
});

it('handles seeded record mutations and truncations without leaking native parser errors', async () => {
  const original = new Uint8Array(streamedEntry(true));
  await fc.assert(
    fc.asyncProperty(
      fc.integer({ min: 0, max: original.length - 1 }),
      fc.integer({ min: 1, max: 255 }),
      fc.boolean(),
      async (position, mask, truncate) => {
        const bytes = original.slice(0, truncate ? position : original.length);
        if (!truncate) bytes[position] ^= mask;
        try {
          await validateZip(bytes.buffer, true);
        } catch (error) {
          expect(error).toBeInstanceOf(Error);
          expect((error as Error).message).toMatch(/^Invalid ZIP:/);
        }
      },
    ),
    { seed: 20261023, numRuns: 300 },
  );
});

it('bounds errors for seeded mutations of a multi-entry workbook archive', async () => {
  const original = new Uint8Array(archive());
  await validateZip(original.buffer, true);
  let rejected = 0;
  await fc.assert(
    fc.asyncProperty(
      fc.integer({ min: 0, max: original.length - 1 }),
      fc.integer({ min: 1, max: 255 }),
      fc.integer({ min: 0, max: 2 }),
      async (position, mask, mode) => {
        const bytes = original.slice(0, mode === 2 ? position : original.length);
        if (mode !== 2) {
          bytes[position] ^= mask;
          if (mode === 1 && position + 1 < bytes.length) bytes[position + 1] ^= mask;
        }
        try {
          await validateZip(bytes.buffer, true);
        } catch (error) {
          rejected++;
          expect(error).toBeInstanceOf(Error);
          expect((error as Error).message).toMatch(/^Invalid ZIP:/);
        }
      },
    ),
    { seed: 20261031, numRuns: 300 },
  );
  expect(rejected).toBeGreaterThan(0);
});
