import { expect, test } from '@playwright/test';

test('rejects damaged gzip trailers in verification, random reads and Worker import', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { BinaryReader, crc32, writeBinary } = await import(
      /* @vite-ignore */ `${root}/binary-file.ts`
    );
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/binary-verify.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision,
      metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const pageBytes = encodeBlock(
      {
        '0:0': {
          rowId: 0,
          columnId: 0,
          input: { type: 'string', value: 'repeated content '.repeat(500) },
        },
      },
      [],
      [],
      {},
    );
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes: pageBytes };
    }
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const original = new Uint8Array(await new Blob(parts).arrayBuffer()),
      view = new DataView(original.buffer);
    const start = 12 + view.getUint32(4, true),
      length = view.getUint32(start + 16, true);
    if (view.getUint32(start + 12, true) !== 1 || length < 18)
      throw new Error('Expected gzip fixture');
    const verified = await verifyBinary(original);
    const codes: string[][] = [];
    for (const trailerByte of [length - 8, length - 4]) {
      const bytes = original.slice();
      bytes[start + 28 + trailerByte] ^= 1;
      new DataView(bytes.buffer).setUint32(
        start + 24,
        crc32(bytes.subarray(start + 28, start + 28 + length)),
        true,
      );
      const errors: string[] = [];
      for (const run of [
        () => verifyBinary(bytes),
        async () => (await BinaryReader.open(bytes)).page(0),
        () => book.importBinary(bytes),
      ]) {
        try {
          await run();
          errors.push('accepted');
        } catch (error) {
          errors.push((error as any).code);
        }
      }
      codes.push(errors);
    }
    const after = book.revision,
      value = await sheet.range('A1').getValues();
    const undo = await book.undo(),
      undone = await sheet.range('A1').getValues();
    const redo = await book.redo(),
      redone = await sheet.range('A1').getValues();
    await window.opensheet.open(book.id);
    const reopened = window.opensheet.getWorkbook();
    const reopenedValue = await reopened.getSheets()[0].range('A1').getValues();
    await reopened.importBinary(original);
    const importedValue = await reopened.getSheets()[0].range('A1').getValues();
    return {
      verified,
      codes,
      revision,
      after,
      value,
      undo,
      undone,
      redo,
      redone,
      reopenedValue,
      importedValue,
    };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result.verified.pages).toBe(1);
  expect(result.verified.cells).toBe(1);
  expect(result.codes).toEqual(Array.from({ length: 2 }, () => Array(3).fill('CORRUPT_STORAGE')));
  expect(result.after).toBe(result.revision);
  expect(result.value).toEqual([[222]]);
  expect(result.undo).toBe(true);
  expect(result.undone).toEqual([[111]]);
  expect(result.redo).toBe(true);
  expect(result.redone).toEqual([[222]]);
  expect(result.reopenedValue).toEqual([[222]]);
  expect(result.importedValue).toEqual([['repeated content '.repeat(500)]]);
});

test('reports corrupted compressed pages without changing the workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { BinaryReader, crc32 } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/binary-verify.ts`);
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision,
      parts: BlobPart[] = [];
    for await (const part of book.streamBinary()) parts.push(part as BlobPart);
    const bytes = new Uint8Array(await new Blob(parts).arrayBuffer()),
      view = new DataView(bytes.buffer);
    const start = 12 + view.getUint32(4, true),
      length = view.getUint32(start + 16, true);
    view.setUint32(start + 12, 1, true);
    bytes.fill(0, start + 28, start + 28 + length);
    view.setUint32(start + 24, crc32(bytes.subarray(start + 28, start + 28 + length)), true);
    const codes: string[] = [];
    for (const run of [
      () => verifyBinary(bytes),
      async () => (await BinaryReader.open(bytes)).page(0),
      () => book.importBinary(bytes),
    ]) {
      try {
        await run();
        codes.push('accepted');
      } catch (error) {
        codes.push((error as any).code);
      }
    }
    const after = book.revision,
      value = await sheet.range('A1').getValues();
    const undo = await book.undo(),
      undone = await sheet.range('A1').getValues();
    const redo = await book.redo(),
      redone = await sheet.range('A1').getValues();
    return { codes, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    codes: Array(3).fill('CORRUPT_STORAGE'),
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});

test('classifies invalid file metadata and directory text without changing data', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { BinaryReader, readBinary, crc32 } = await import(
      /* @vite-ignore */ `${root}/binary-file.ts`
    );
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const parts: BlobPart[] = [];
    for await (const part of book.streamBinary()) parts.push(part as BlobPart);
    const original = new Uint8Array(await new Blob(parts).arrayBuffer());
    const originalView = new DataView(original.buffer);
    const metadataLength = originalView.getUint32(4, true);
    const footer = original.length - 24;
    const directoryOffset = Number(originalView.getBigUint64(footer + 4, true));
    const directoryLength = Number(originalView.getBigUint64(footer + 12, true));
    const codes: string[][] = [];
    const inspect = async (damaged: Uint8Array) => {
      const errors: string[] = [];
      for (const run of [
        () => BinaryReader.open(damaged),
        async () => {
          for await (const _ of readBinary(damaged)) {
            /* consume */
          }
        },
        () => book.importBinary(damaged),
      ]) {
        try {
          await run();
          errors.push('accepted');
        } catch (error) {
          errors.push((error as any).code);
        }
      }
      codes.push(errors);
    };
    for (const target of ['metadata', 'directory']) {
      for (const invalid of [0x21, 0xff]) {
        const damaged = original.slice();
        const view = new DataView(damaged.buffer);
        const start = target === 'metadata' ? 12 : directoryOffset;
        const length = target === 'metadata' ? metadataLength : directoryLength;
        damaged[start] = invalid;
        view.setUint32(
          target === 'metadata' ? 8 : footer + 20,
          crc32(damaged.subarray(start, start + length)),
          true,
        );
        await inspect(damaged);
      }
    }
    const nullEntry = original.slice();
    nullEntry.fill(0x20, directoryOffset, directoryOffset + directoryLength);
    nullEntry.set(new TextEncoder().encode('[null]'), directoryOffset);
    new DataView(nullEntry.buffer).setUint32(
      footer + 20,
      crc32(nullEntry.subarray(directoryOffset, directoryOffset + directoryLength)),
      true,
    );
    await inspect(nullEntry);
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { codes, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    codes: Array.from({ length: 5 }, () => [
      'INVALID_ARGUMENT',
      'INVALID_ARGUMENT',
      'INVALID_ARGUMENT',
    ]),
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});

test('rejects invalid binary page encoding without changing workbook history', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const parts: BlobPart[] = [];
    for await (const part of book.streamBinary()) parts.push(part as BlobPart);
    const original = new Uint8Array(await new Blob(parts).arrayBuffer());
    const start = 12 + new DataView(original.buffer).getUint32(4, true);
    const codes: string[] = [];
    for (const [codec, decodedLength] of [
      [0, 0],
      [1, 32 * 1024 * 1024 + 1],
      [2, 1],
      [0, 1],
    ]) {
      const bytes = original.slice();
      const view = new DataView(bytes.buffer);
      view.setUint32(start + 12, codec, true);
      view.setUint32(start + 20, decodedLength, true);
      try {
        await book.importBinary(bytes);
        codes.push('accepted');
      } catch (error) {
        codes.push((error as { code: string }).code);
      }
    }
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { codes, revision, after, value, undo, undone, redo, redone };
  });
  expect(result).toEqual({
    codes: ['LIMIT_EXCEEDED', 'LIMIT_EXCEEDED', 'INVALID_ARGUMENT', 'INVALID_ARGUMENT'],
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});
