import { type WorkbookSnapshot } from '../packages/core/src/types.js';
import { encodeWorkbookFile, decodeWorkbookFile } from '../packages/core/src/file-codec.js';
import { createSnapshot } from '../packages/core/src/model.js';
import { afterEach, describe, expect, it, vi } from 'vitest';
import fc from 'fast-check';
import {
  createWorkbook,
  openWorkbook,
  Workbook,
  workbookFileSchema,
  key,
  type WorkbookFile,
} from '@opensheetjs/core';
import { encodeAxis } from '../packages/core/src/file-codec.js';
import { encodeBlock, decodeBlock } from '../packages/core/src/block-codec.js';
import { BinaryReader } from '../packages/core/src/binary-file.js';
import { parseWorkbookJSON } from '../packages/core/src/json-stream.js';
const books: Workbook[] = [];
afterEach(() => books.splice(0).forEach((book) => book.dispose()));
const track = async (snapshot?: WorkbookSnapshot) => {
  const book = snapshot
    ? await new Workbook(encodeWorkbookFile(snapshot)).ready()
    : await createWorkbook();
  books.push(book);
  return book;
};
const streamFile = async (book: Workbook) => {
  let text = '';
  for await (const chunk of book.streamJSON()) text += chunk;
  return JSON.parse(text) as WorkbookFile;
};
const binary = async (book: Workbook) => {
  const parts: Uint8Array[] = [];
  for await (const bytes of book.streamBinary()) parts.push(bytes);
  return new Blob(parts as BlobPart[]);
};
const canonical = (file: WorkbookFile) => ({ ...file, revision: 0, workbookId: 'ignored' });
describe('native column files and binary pages', () => {
  it('preserves every type, sparse property, negative zero and Unicode through JSON and binary', async () => {
    const snapshot = createSnapshot({
      sheets: [{ name: 'One', rows: 130, columns: 40 }, { name: 'Two' }],
    });
    snapshot.sheetOrder.reverse();
    snapshot.dateSystem = '1904';
    snapshot.extensions = { test: { value: 0 } };
    snapshot.styles.s = { bold: true, numberFormat: '0%' };
    const sheet = snapshot.sheets[0];
    sheet.rowOrder.splice(5, 0, sheet.nextRowId++);
    sheet.rowOrder.reverse();
    sheet.rows[sheet.rowOrder[9]] = { hidden: true, size: 31 };
    sheet.columns[2] = { size: 111 };
    sheet.freeze = { rows: 1, columns: 2 };
    const inputs = [
      { type: 'blank' },
      { type: 'string', value: '' },
      { type: 'string', value: '=A1\ud800你好😀' },
      { type: 'number', value: -0 },
      { type: 'boolean', value: false },
      { type: 'formula', expression: 'SUM(A1:A2)' },
      { type: 'error', code: '#REF!' },
    ] as const;
    inputs.forEach((input, i) => {
      const rowId = sheet.rowOrder[i * 18],
        columnId = sheet.columnOrder[i * 5];
      sheet.cells[key(rowId, columnId)] = {
        rowId,
        columnId,
        input,
        styleId: 's',
        note: '注释',
        link: { target: 'https://example.com', tooltip: 'link' },
        numberFormat: '0.00',
        cached: { type: 'number', value: -0 },
      };
    });
    const file = encodeWorkbookFile(snapshot);
    expect(workbookFileSchema.safeParse(file).success).toBe(true);
    expect(encodeWorkbookFile(decodeWorkbookFile(JSON.parse(JSON.stringify(file))))).toEqual(file);
    const book = await track(snapshot);
    expect(await streamFile(book)).toEqual(await book.toJSON());
    const restored = await track();
    await restored.importBinary(await binary(book));
    expect(canonical(await restored.toJSON())).toEqual(canonical(await book.toJSON()));
    const cell = await book.getCell(
      sheet.id,
      sheet.rowOrder.indexOf(sheet.rowOrder[3 * 18]),
      sheet.columnOrder.indexOf(sheet.columnOrder[3 * 5]),
    );
    expect(Object.is(cell!.input.type === 'number' ? cell!.input.value : 1, -0)).toBe(true);
  });
  it('round-trips arbitrary order and compresses forward/reverse segments', () => {
    fc.assert(
      fc.property(
        fc.uniqueArray(fc.integer({ min: 0, max: 500 }), { minLength: 1, maxLength: 100 }),
        (order) => {
          const snapshot = createSnapshot();
          snapshot.sheets[0].rowOrder = order;
          snapshot.sheets[0].nextRowId = 501;
          const file = encodeWorkbookFile(snapshot);
          expect(encodeWorkbookFile(decodeWorkbookFile(file))).toEqual(file);
        },
      ),
    );
    expect(encodeAxis([3, 2, 1, 5, 6], 7).order).toEqual([
      [3, 3, -1],
      [5, 2, 1],
    ]);
  });
  it('rejects malformed lengths, duplicate cells/pages, old schemas and unsafe objects', () => {
    const snapshot = createSnapshot();
    snapshot.sheets[0].cells[key(0, 0)] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'number', value: 0 },
    };
    const file = encodeWorkbookFile(snapshot);
    const mutations: Array<(f: any) => void> = [
      (f) => (f.schemaVersion = 2),
      (f) => (f.sheets[0].rows.order = [[0, 100001, 1]]),
      (f) => f.sheets[0].blocks.push(f.sheets[0].blocks[0]),
      (f) => f.sheets[0].blocks[0].data.numbers.positions.push(0),
      (f) => (f.sheets[0].blocks[0].data.numbers.values[0] = 'bad'),
      (f) => (f.sheets[0].blocks[0].data.styles = [[0, 1, 99]]),
      (f) => (f.sheets[0].blocks[0].row = 100),
      (f) => (f.sheets[0].rows.meta = { '1000': {} }),
      (f) => f.sheets.push(f.sheets[0]),
    ];
    for (const mutate of mutations) {
      const copy = structuredClone(file);
      mutate(copy);
      expect(() => decodeWorkbookFile(copy)).toThrow();
    }
    expect(() => decodeWorkbookFile(JSON.parse('{"__proto__":{}}'))).toThrow();
  });
  it('exports sparse distant cells by occupied pages and expires stale imported caches', async () => {
    const snapshot = createSnapshot({ sheets: [{ name: 'Sparse', rows: 100000, columns: 1000 }] });
    const sh = snapshot.sheets[0];
    sh.cells[key(0, 0)] = {
      rowId: 0,
      columnId: 0,
      input: { type: 'formula', expression: '1+1' },
      cached: { type: 'number', value: 2 },
    };
    sh.cells[key(99999, 999)] = {
      rowId: 99999,
      columnId: 999,
      input: { type: 'boolean', value: false },
    };
    const book = await track(snapshot),
      read = vi.spyOn(book, 'streamRange'),
      file = await streamFile(book);
    expect(read).not.toHaveBeenCalled();
    expect(file.sheets[0].blocks).toHaveLength(2);
    expect(JSON.stringify(file).length).toBeLessThan(1700);
    await book
      .getSheets()[0]
      .range('B1')
      .setValues([[1]]);
    expect(
      decodeWorkbookFile(await book.toJSON()).sheets[0].cells[key(0, 0)].cached,
    ).toBeUndefined();
  });
  it('imports chunks with reordered metadata and preserves the original on a late error', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[7]]);
    const file = await book.toJSON();
    const reordered = {
      sheets: file.sheets,
      styles: file.styles,
      ...Object.fromEntries(
        Object.entries(file).filter(([key]) => key !== 'sheets' && key !== 'styles'),
      ),
    };
    const text = JSON.stringify(reordered);
    const bytes = new TextEncoder().encode(text);
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        for (let i = 0; i < bytes.length; i += 3) controller.enqueue(bytes.slice(i, i + 3));
        controller.close();
      },
    });
    await book.importJSON(stream);
    expect(await book.getSheets()[0].range('A1').getValues()).toEqual([[7]]);
    const before = await book.toJSON();
    await expect(book.importJSON(text.slice(0, -1))).rejects.toThrow();
    expect(await book.toJSON()).toEqual(before);
    await expect(
      book.importJSON(text.replace('"schemaVersion":3', '"schemaVersion":3,"schemaVersion":3')),
    ).rejects.toThrow(/Duplicate/);
    for (const text of ['{"x":1,}', '{"x":[1,]}', '{}{}', '{"x":01}']) {
      await expect(async () => {
        for await (const _ of parseWorkbookJSON(text)) {
        }
      }).rejects.toThrow();
    }
  });
  it('allows independent binary page reads and rejects corrupted input atomically', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1:B1')
      .setValues([[0, false]]);
    const blob = await binary(book);
    const reader = await BinaryReader.open(blob);
    expect(Object.values(decodeBlock(await reader.page(0)).cells).map((c) => c.input)).toEqual([
      { type: 'number', value: 0 },
      { type: 'boolean', value: false },
    ]);
    const bytes = new Uint8Array(await blob.arrayBuffer());
    bytes[reader.directory[0].offset + 28] ^= 1;
    const before = await book.toJSON();
    await expect(book.importBinary(bytes)).rejects.toThrow(/checksum/i);
    expect(await book.toJSON()).toEqual(before);
  });
  it('preserves identities through insertion, undo and redo without rewriting unrelated data pages', async () => {
    const book = await track();
    await book
      .getSheets()[0]
      .range('A1:B2')
      .setValues([
        [0, false],
        ['', 'x'],
      ]);
    const before = await book.toJSON();
    await book.getSheets()[0].insertRows(0);
    await book.getSheets()[0].insertColumns(1);
    const after = await book.toJSON();
    await book.undo();
    await book.undo();
    expect((await book.toJSON()).sheets).toEqual(before.sheets);
    await book.redo();
    await book.redo();
    expect((await book.toJSON()).sheets).toEqual(after.sheets);
  });
  it('aborts exports and rejects mixed revisions', async () => {
    const book = await track();
    const abort = new AbortController();
    abort.abort();
    await expect(book.streamJSON({ signal: abort.signal }).next()).rejects.toMatchObject({
      code: 'ABORTED',
    });
    const stream = book.streamJSON();
    await stream.next();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[1]]);
    await expect(stream.next()).rejects.toMatchObject({ code: 'REVISION_CONFLICT' });
  });
  it('retains runtime style identities through metadata persistence and reopening', async () => {
    const database = `styles-${crypto.randomUUID()}`;
    const book = await createWorkbook({}, { database });
    books.push(book);
    const sheetId = book.getSheets()[0].id;
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[0.25]]);
    await book.getSheets()[0].range('A1').setStyle({ bold: true, numberFormat: '0%' });
    const before = await book.toJSON();
    await book.close();
    const reopened = await openWorkbook(book.id, { database });
    books.push(reopened);
    const cell = await reopened.getCell(sheetId, 0, 0);
    expect(reopened.getStyle(cell?.styleId)).toEqual({ bold: true, numberFormat: '0%' });
    expect(await reopened.toJSON()).toEqual(before);
    expect(await streamFile(reopened)).toEqual(before);
    const restored = await track();
    await restored.importBinary(await binary(reopened));
    expect(canonical(await restored.toJSON())).toEqual(canonical(before));
  });
  it('omits unused styles consistently across complete JSON, streaming JSON and binary', async () => {
    const snapshot = createSnapshot();
    snapshot.styles.unused = { italic: true };
    snapshot.styles.used = { bold: true };
    const sh = snapshot.sheets[0];
    sh.cells[key(0, 0)] = { rowId: 0, columnId: 0, input: { type: 'blank' }, styleId: 'used' };
    const book = await track(snapshot);
    const file = await book.toJSON();
    expect(file.styles).toEqual([{ bold: true }]);
    expect(await streamFile(book)).toEqual(file);
    const restored = await track();
    await restored.importBinary(await binary(book));
    expect(canonical(await restored.toJSON())).toEqual(canonical(file));
    await book.getSheets()[0].range('A1').clear({ all: true });
    expect((await book.toJSON()).styles).toEqual([]);
    await book.undo();
    expect((await book.toJSON()).styles).toEqual([{ bold: true }]);
  });
  it('keeps repeated styles and text compact in binary pages', () => {
    const cells: Record<string, import('@opensheetjs/core').CellRecord> = {};
    for (let i = 0; i < 64; i++)
      cells[key(i, 0)] = {
        rowId: i,
        columnId: 0,
        input: { type: 'string', value: 'repeated' },
        styleId: 'style_deadbeef',
      };
    const encoded = encodeBlock(cells, [], [], {});
    expect(encoded).toBeInstanceOf(Uint8Array);
    expect(encoded.length).toBeLessThan(1500);
    expect(decodeBlock(encoded).cells).toEqual(cells);
  });
});
