import { expect, test } from '@playwright/test';
import { resolve } from 'node:path';

test('binary import agrees with random reader on reordered and unknown directory fields', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(
    async ({ url, verifyUrl }) => {
      const { BinaryReader, readBinary, crc32 } = await import(/* @vite-ignore */ url);
      const { verifyBinary } = await import(/* @vite-ignore */ verifyUrl);
      const app = window.opensheet;
      let book = app.getWorkbook();
      let sheet = book.getSheets()[0];
      await sheet.range('A1').setValues([[5]]);
      await sheet.range('A70').setValues([[9]]);
      const parts: Uint8Array[] = [];
      for await (const chunk of book.streamBinary()) parts.push(chunk);
      const source = new Uint8Array(await new Blob(parts as BlobPart[]).arrayBuffer());
      const rewrite = (transform: (entry: any) => unknown) => {
        const view = new DataView(source.buffer),
          footer = source.length - 24;
        const offset = Number(view.getBigUint64(footer + 4, true)),
          length = Number(view.getBigUint64(footer + 12, true));
        const entries = JSON.parse(
          new TextDecoder().decode(source.subarray(offset, offset + length)),
        );
        const directory = new TextEncoder().encode(JSON.stringify(entries.map(transform), null, 2));
        const output = new Uint8Array(offset + directory.length + 24),
          next = new DataView(output.buffer);
        output.set(source.subarray(0, offset));
        output.set(directory, offset);
        next.setUint32(offset - 4, directory.length, true);
        const end = offset + directory.length;
        output.set([79, 83, 66, 69], end);
        next.setBigUint64(end + 4, BigInt(offset), true);
        next.setBigUint64(end + 12, BigInt(directory.length), true);
        next.setUint32(end + 20, crc32(directory), true);
        return output;
      };
      const reordered = rewrite((e) => ({
        length: e.length,
        offset: e.offset,
        column: e.column,
        row: e.row,
        sheet: e.sheet,
      }));
      const verified = await verifyBinary(reordered);
      const random = await BinaryReader.open(reordered);
      let pages = 0;
      let samePages = true;
      for await (const part of readBinary(reordered)) {
        if (!('bytes' in part)) continue;
        const bytes = await random.page(pages++);
        if (
          bytes.length !== part.bytes.length ||
          bytes.some((byte: number, index: number) => byte !== part.bytes[index])
        )
          samePages = false;
      }
      await sheet.range('A1').setValues([[123]]);
      await book.importBinary(reordered);
      sheet = book.getSheets()[0];
      const imported = [await sheet.range('A1').getValues(), await sheet.range('A70').getValues()];
      await sheet.range('A1').setValues([[111]]);
      const revision = book.revision;
      const invalid = rewrite((e) => ({ ...e, unknown: true }));
      const randomError = await BinaryReader.open(invalid).then(
        () => 'success',
        (e: { code?: string }) => e.code,
      );
      const verifyError = await verifyBinary(invalid).then(
        () => 'success',
        (e: { code?: string }) => e.code,
      );
      const importError = await book.importBinary(invalid).then(
        () => 'success',
        (e: { code?: string }) => e.code,
      );
      const afterFailure = { revision: book.revision, value: await sheet.range('A1').getValues() };
      await book.undo();
      const undo = await sheet.range('A1').getValues();
      await book.redo();
      await app.open(book.id);
      book = app.getWorkbook();
      sheet = book.getSheets()[0];
      return {
        directory: random.directory.length,
        pages,
        samePages,
        verified,
        imported,
        revision,
        randomError,
        verifyError,
        importError,
        afterFailure,
        undo,
        reopened: [await sheet.range('A1').getValues(), await sheet.range('A70').getValues()],
      };
    },
    {
      url: `/@fs/${resolve('packages/core/src/binary-file.ts')}`,
      verifyUrl: `/@fs/${resolve('packages/core/src/binary-verify.ts')}`,
    },
  );
  expect(result.directory).toBeGreaterThanOrEqual(2);
  expect(result.pages).toBe(result.directory);
  expect(result.samePages).toBe(true);
  expect(result.verified.pages).toBe(result.directory);
  expect(result.imported).toEqual([[[5]], [[9]]]);
  expect(result.randomError).toBe('INVALID_ARGUMENT');
  expect(result.verifyError).toBe('INVALID_ARGUMENT');
  expect(result.importError).toBe('INVALID_ARGUMENT');
  expect(result.afterFailure).toEqual({ revision: result.revision, value: [[111]] });
  expect(result.undo).toEqual([[5]]);
  expect(result.reopened).toEqual([[[111]], [[9]]]);
});
