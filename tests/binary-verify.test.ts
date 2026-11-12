import { expect, it, vi } from 'vitest';
import { createWorkbook, verifyBinary } from '@opensheetjs/core';
import { Database } from '../packages/core/src/storage.js';
import { encodeBlock } from '../packages/core/src/block-codec.js';
import { BinaryReader, writeBinary } from '../packages/core/src/binary-file.js';

it('verifies every page without writes and rejects corruption or cancellation', async () => {
  const book = await createWorkbook({
    sheets: [
      { name: 'First', rows: 130, columns: 2 },
      { name: 'Second', rows: 10, columns: 2 },
    ],
  });
  try {
    const [first, second] = book.getSheets();
    await first.range('A1').setValues([[1]]);
    await first.range('A65').setValues([['text']]);
    await second.range('A1').setValues([[true]]);
    const parts: Uint8Array[] = [];
    for await (const chunk of book.streamBinary()) parts.push(chunk);
    const bytes = new Uint8Array(await new Blob(parts as BlobPart[]).arrayBuffer());
    const write = vi.spyOn(Database.prototype, 'write');
    try {
      expect(await verifyBinary(bytes)).toEqual({
        workbookId: book.id,
        sheets: 2,
        pages: 3,
        cells: 3,
      });
      expect(write).not.toHaveBeenCalled();
      const reader = await BinaryReader.open(bytes);
      const corrupted = bytes.slice();
      corrupted[reader.directory[1].offset + 28] ^= 1;
      await expect(verifyBinary(corrupted)).rejects.toThrow();
      const controller = new AbortController();
      controller.abort();
      await expect(verifyBinary(bytes, { signal: controller.signal })).rejects.toMatchObject({
        name: 'AbortError',
      });
      expect(write).not.toHaveBeenCalled();
    } finally {
      write.mockRestore();
    }
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('rejects page cells outside sheet axes even when the binary framing is valid', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 10, columns: 2 }] });
  try {
    const metadata = await book.toJSON();
    const bytes = encodeBlock(
      { '20:0': { rowId: 20, columnId: 0, input: { type: 'number', value: 1 } } },
      [20],
      [0],
      {},
    );
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes };
    }
    const parts: Uint8Array[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts as BlobPart[]);
    await BinaryReader.open(file);
    await expect(verifyBinary(file)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('rejects page styles absent from workbook metadata despite valid page framing', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 10, columns: 2 }] });
  try {
    const metadata = await book.toJSON();
    const bytes = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 1 }, styleId: 'missing' } },
      [0],
      [0],
      {},
    );
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes };
    }
    const parts: Uint8Array[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts as BlobPart[]);
    await expect(verifyBinary(file)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
  } finally {
    await book.deleteStorage();
  }
});

it('rejects inline metadata blocks rather than reporting an incomplete binary backup as valid', async () => {
  const book = await createWorkbook();
  try {
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[91]]);
    const metadata = await book.toJSON();
    expect(metadata.sheets[0].blocks).toHaveLength(1);
    async function* pages(): AsyncGenerator<{
      sheet: number;
      row: number;
      column: number;
      bytes: Uint8Array;
    }> {}
    const parts: Uint8Array[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts as BlobPart[]);
    await expect(BinaryReader.open(file)).rejects.toThrow('Binary metadata contains inline blocks');
    await expect(verifyBinary(file)).rejects.toThrow('Binary metadata contains inline blocks');
    const revision = book.revision;
    await expect(book.importBinary(file)).rejects.toThrow('Binary metadata contains inline blocks');
    expect(book.revision).toBe(revision);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[91]]);
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});

it('rejects ambiguous and unrecognized page metadata without replacing the workbook', async () => {
  const book = await createWorkbook();
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[91]]);
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry) => (entry.blocks = []));
    const encoded = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } } },
      [0],
      [0],
      {},
    );
    const originalLength = new DataView(encoded.buffer, encoded.byteOffset).getUint32(4, true);
    const oldBase = Math.ceil((8 + originalLength) / 8) * 8;
    const original = JSON.parse(new TextDecoder().decode(encoded.subarray(8, 8 + originalLength)));
    for (const change of [
      (meta: any) => {
        meta.data.numbers = { positions: [0], values: [999] };
      },
      (meta: any) => {
        meta.mandatoryFutureField = true;
      },
      (meta: any) => {
        meta.sections[0].mandatoryFutureField = true;
      },
      (meta: any) => {
        meta.sections[0].positions.mandatoryFutureField = true;
      },
    ]) {
      const meta = structuredClone(original);
      change(meta);
      const header = new TextEncoder().encode(JSON.stringify(meta));
      const base = Math.ceil((8 + header.length) / 8) * 8;
      const page = new Uint8Array(base + encoded.length - oldBase);
      page.set(encoded.subarray(0, 4));
      new DataView(page.buffer).setUint32(4, header.length, true);
      page.set(header, 8);
      page.set(encoded.subarray(oldBase), base);
      async function* pages() {
        yield { sheet: 0, row: 0, column: 0, bytes: page };
      }
      const parts: Uint8Array[] = [];
      for await (const part of writeBinary(metadata, pages())) parts.push(part);
      const file = new Blob(parts as BlobPart[]);
      await expect(verifyBinary(file)).rejects.toMatchObject({ code: 'CORRUPT_STORAGE' });
      const revision = book.revision;
      await expect(book.importBinary(file)).rejects.toMatchObject({ code: 'CORRUPT_STORAGE' });
      expect(book.revision).toBe(revision);
      expect(await sheet.range('A1').getValues()).toEqual([[91]]);
    }
  } finally {
    await book.deleteStorage();
    await book.close();
  }
});
