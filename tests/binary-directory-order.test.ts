import { expect, it } from 'vitest';
import { createWorkbook } from '@opensheetjs/core';
import { BinaryReader, crc32, readBinary } from '../packages/core/src/binary-file.js';

function rewriteDirectory(source: Uint8Array, transform: (entry: any) => unknown) {
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
  const footer = source.length - 24;
  const offset = Number(view.getBigUint64(footer + 4, true));
  const length = Number(view.getBigUint64(footer + 12, true));
  const original = JSON.parse(new TextDecoder().decode(source.subarray(offset, offset + length)));
  const directory = new TextEncoder().encode(JSON.stringify(original.map(transform), null, 2));
  const output = new Uint8Array(offset + directory.length + 24);
  output.set(source.subarray(0, offset));
  output.set(directory, offset);
  const result = new DataView(output.buffer);
  result.setUint32(offset - 4, directory.length, true);
  const newFooter = offset + directory.length;
  output.set([79, 83, 66, 69], newFooter);
  result.setBigUint64(newFooter + 4, BigInt(offset), true);
  result.setBigUint64(newFooter + 12, BigInt(directory.length), true);
  result.setUint32(newFooter + 20, crc32(directory), true);
  return output;
}

it('accepts reordered directory keys and rejects unknown keys consistently', async () => {
  const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 100, columns: 2 }] });
  try {
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[5]]);
    await sheet.range('A70').setValues([[9]]);
    const chunks: Uint8Array[] = [];
    for await (const chunk of book.streamBinary()) chunks.push(chunk);
    const source = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
    const reordered = rewriteDirectory(source, (e) => ({
      length: e.length,
      offset: e.offset,
      column: e.column,
      row: e.row,
      sheet: e.sheet,
    }));
    const random = await BinaryReader.open(reordered);
    expect(random.directory).toHaveLength(2);
    const pages = [];
    for await (const part of readBinary(reordered)) if ('bytes' in part) pages.push(part);
    expect(pages).toHaveLength(2);
    for (let i = 0; i < pages.length; i++) expect(await random.page(i)).toEqual(pages[i].bytes);
    for (const transform of [
      (e: any) => ({ ...e, unknown: true }),
      ({ length: _length, ...rest }: any) => rest,
      (e: any) => ({ ...e, row: e.row + 1 }),
    ]) {
      const invalid = rewriteDirectory(source, transform);
      const reader = await BinaryReader.open(invalid).catch((error) => error);
      if (reader instanceof BinaryReader) {
        // Random access checks the frame lazily, so mismatched coordinates fail on page read.
        await expect(reader.page(0)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
      } else expect(reader).toMatchObject({ code: 'INVALID_ARGUMENT' });
      await expect(async () => {
        for await (const _ of readBinary(invalid)) {
        }
      }).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    }
  } finally {
    await book.close();
  }
});
