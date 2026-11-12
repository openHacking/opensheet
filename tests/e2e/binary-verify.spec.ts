import { expect, test } from '@playwright/test';

test('checks large random-access page headers before reading payload', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { BinaryReader, writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const book = window.opensheet.getWorkbook();
    const metadata = await book.toJSON();
    metadata.sheets.forEach((sheet: any) => (sheet.blocks = []));
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
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const original = new Uint8Array(await new Blob(parts).arrayBuffer());
    const offset = 12 + new DataView(original.buffer).getUint32(4, true);
    const valid = await BinaryReader.open(original);
    const read = await valid.page(0);
    const matches =
      read.length === payload.length &&
      read.every((value: number, i: number) => value === payload[i]);
    const damaged = original.slice();
    new DataView(damaged.buffer).setUint32(offset + 20, 32 * 1024 * 1024 + 1, true);
    const blob = new Blob([damaged]);
    const reader = await BinaryReader.open(blob);
    const originalSlice = blob.slice.bind(blob);
    const sizes: number[] = [];
    blob.slice = (start = 0, end = blob.size, type) => {
      sizes.push(end - start);
      return originalSlice(start, end, type);
    };
    let code = '';
    try {
      await reader.page(0);
    } catch (error) {
      code = (error as any).code;
    }
    return { matches, code, sizes, encodedLength: reader.directory[0].length };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result.matches).toBe(true);
  expect(result.encodedLength).toBeGreaterThan(64 * 1024);
  expect(result.code).toBe('LIMIT_EXCEEDED');
  expect(result.sizes).toEqual([28]);
});

test('verifies native backup pages without replacing the current workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (moduleUrl) => {
    const { verifyBinary, BinaryReader } = await import(/* @vite-ignore */ moduleUrl);
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[987]]);
    const revision = book.revision,
      parts: BlobPart[] = [];
    for await (const part of book.streamBinary()) parts.push(part as BlobPart);
    const blob = new Blob(parts);
    const verified = await verifyBinary(blob);
    const reader = await BinaryReader.open(blob);
    const bytes = new Uint8Array(await blob.arrayBuffer());
    bytes[reader.directory[0].offset + 28] ^= 1;
    let rejected = false;
    try {
      await verifyBinary(bytes);
    } catch {
      rejected = true;
    }
    return {
      verified,
      rejected,
      before: revision,
      after: book.revision,
      value: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result.verified.workbookId).toBeTruthy();
  expect(result.verified.pages).toBeGreaterThan(0);
  expect(result.verified.cells).toBeGreaterThan(0);
  expect(result.rejected).toBe(true);
  expect(result.after).toBe(result.before);
  expect(result.value).toEqual([[987]]);
});

test('rejects ambiguous binary pages without changing data or history', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/index.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const bytes = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } } },
      [0],
      [0],
      {},
    );
    const length = new DataView(bytes.buffer, bytes.byteOffset).getUint32(4, true);
    const oldBase = Math.ceil((8 + length) / 8) * 8;
    const meta = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + length)));
    meta.data.numbers = { positions: [0], values: [999] };
    const header = new TextEncoder().encode(JSON.stringify(meta));
    const base = Math.ceil((8 + header.length) / 8) * 8;
    const altered = new Uint8Array(base + bytes.length - oldBase);
    altered.set(bytes.subarray(0, 4));
    new DataView(altered.buffer).setUint32(4, header.length, true);
    altered.set(header, 8);
    altered.set(bytes.subarray(oldBase), base);
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes: altered };
    }
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts);
    let verifyCode = '',
      importCode = '';
    try {
      await verifyBinary(file);
    } catch (error) {
      verifyCode = (error as any).code;
    }
    try {
      await book.importBinary(file);
    } catch (error) {
      importCode = (error as any).code;
    }
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { verifyCode, importCode, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    verifyCode: 'CORRUPT_STORAGE',
    importCode: 'CORRUPT_STORAGE',
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});

test('rejects an oversized position run without replacing the workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/index.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const cells: Record<string, unknown> = {};
    for (let column = 0; column < 32; column++)
      cells[`0:${column}`] = {
        rowId: 0,
        columnId: column,
        input: { type: 'number', value: column },
      };
    const bytes = encodeBlock(cells, [], [], {});
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const length = view.getUint32(4, true);
    const meta = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + length)));
    if (meta.sections[0].positions.mode !== 1) throw new Error('Fixture is not run-encoded');
    const base = Math.ceil((8 + length) / 8) * 8;
    view.setUint16(base + meta.sections[0].positions.offset + 2, 2048, true);
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes };
    }
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts);
    let verifyCode = '',
      importCode = '';
    try {
      await verifyBinary(file);
    } catch (error) {
      verifyCode = (error as any).code;
    }
    try {
      await book.importBinary(file);
    } catch (error) {
      importCode = (error as any).code;
    }
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { verifyCode, importCode, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    verifyCode: 'CORRUPT_STORAGE',
    importCode: 'CORRUPT_STORAGE',
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});

test('rejects amplified dictionary text without replacing the workbook', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/index.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const cells: Record<string, unknown> = {};
    for (let i = 0; i < 400; i++) {
      const row = Math.floor(i / 32),
        column = i % 32;
      cells[`${row}:${column}`] = {
        rowId: row,
        columnId: column,
        input: { type: 'string', value: 'x'.repeat(32768) },
      };
    }
    const bytes = encodeBlock(cells, [], [], {});
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes };
    }
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts);
    let verifyCode = '',
      importCode = '';
    try {
      await verifyBinary(file);
    } catch (error) {
      verifyCode = (error as any).code;
    }
    try {
      await book.importBinary(file);
    } catch (error) {
      importCode = (error as any).code;
    }
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { verifyCode, importCode, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    verifyCode: 'LIMIT_EXCEEDED',
    importCode: 'LIMIT_EXCEEDED',
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});

test('rejects fractional position counts atomically', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/index.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const bytes = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 1 } } },
      [],
      [],
      {},
    );
    const length = new DataView(bytes.buffer, bytes.byteOffset).getUint32(4, true);
    const oldBase = Math.ceil((8 + length) / 8) * 8;
    const meta = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + length)));
    if (meta.sections[0].positions.mode !== 0) throw new Error('Fixture is not list-encoded');
    meta.sections[0].positions.count = 0.5;
    meta.sections[0].positions.bytes = 1;
    const header = new TextEncoder().encode(JSON.stringify(meta));
    const base = Math.ceil((8 + header.length) / 8) * 8;
    const altered = new Uint8Array(base + bytes.length - oldBase);
    altered.set(bytes.subarray(0, 4));
    new DataView(altered.buffer).setUint32(4, header.length, true);
    altered.set(header, 8);
    altered.set(bytes.subarray(oldBase), base);
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes: altered };
    }
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts);
    let verifyCode = '',
      importCode = '';
    try {
      await verifyBinary(file);
    } catch (error) {
      verifyCode = (error as any).code;
    }
    try {
      await book.importBinary(file);
    } catch (error) {
      importCode = (error as any).code;
    }
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { verifyCode, importCode, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    verifyCode: 'CORRUPT_STORAGE',
    importCode: 'CORRUPT_STORAGE',
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});

test('classifies malformed page JSON and UTF-8 while preserving workbook history', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/index.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const original = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'string', value: 'hello' } } },
      [],
      [],
      {},
    );
    const length = new DataView(original.buffer, original.byteOffset).getUint32(4, true);
    const meta = JSON.parse(new TextDecoder().decode(original.subarray(8, 8 + length)));
    if (meta.sections[0].type !== 0) throw new Error('Fixture is not plain UTF-8');
    const textOffset = Math.ceil((8 + length) / 8) * 8 + meta.sections[0].offset + 12;
    const codes: string[][] = [];
    for (const [offset, value] of [
      [8, 0x21],
      [8, 0xff],
      [textOffset, 0xff],
    ]) {
      const bytes = original.slice();
      bytes[offset] = value;
      async function* pages() {
        yield { sheet: 0, row: 0, column: 0, bytes };
      }
      const parts: BlobPart[] = [];
      for await (const part of writeBinary(metadata, pages())) parts.push(part);
      const file = new Blob(parts);
      let verifyCode = '',
        importCode = '';
      try {
        await verifyBinary(file);
      } catch (error) {
        verifyCode = (error as any).code;
      }
      try {
        await book.importBinary(file);
      } catch (error) {
        importCode = (error as any).code;
      }
      codes.push([verifyCode, importCode]);
    }
    const after = book.revision;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    return { codes, revision, after, value, undo, undone, redo, redone };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    codes: Array.from({ length: 3 }, () => ['CORRUPT_STORAGE', 'CORRUPT_STORAGE']),
    revision: result.revision,
    after: result.revision,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
  });
});
