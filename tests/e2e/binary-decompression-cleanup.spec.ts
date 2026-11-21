import { expect, test } from '@playwright/test';

test('a stalled decompressor cancellation does not hide binary import overflow', async ({
  page,
}) => {
  test.setTimeout(30_000);
  await page.route('**/engine.worker.js*', async (route) => {
    const response = await route.fetch();
    await route.fulfill({
      response,
      body:
        (await response.text()) +
        `
let stallDecompressionCancel = false;
addEventListener('message', ({data}) => {
  if (typeof data.stallDecompressionCancel === 'boolean') stallDecompressionCancel = data.stallDecompressionCancel;
});
const NativeDecompression = globalThis.DecompressionStream;
globalThis.DecompressionStream = class extends NativeDecompression {
  constructor(format) {
    super(format);
    const getReader = this.readable.getReader.bind(this.readable);
    this.readable.getReader = (...args) => {
      const reader = getReader(...args);
      const cancel = reader.cancel.bind(reader);
      reader.cancel = (...args) => {
        if (!stallDecompressionCancel) return cancel(...args);
        postMessage({decompressionCancelStalled:true});
        return new Promise(() => {});
      };
      return reader;
    };
  }
};`,
    });
  });
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    let book = window.opensheet.getWorkbook();
    let sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision;
    const bytesBefore = (await book.storageStats()).bytes;
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const pageBytes = encodeBlock(
      {
        '0:0': {
          rowId: 0,
          columnId: 0,
          input: { type: 'string', value: 'repeated '.repeat(1000) },
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
    const valid = new Uint8Array(await new Blob(parts).arrayBuffer());
    const bad = valid.slice();
    const view = new DataView(bad.buffer);
    const start = 12 + view.getUint32(4, true);
    if (view.getUint32(start + 12, true) !== 1) throw new Error('Expected compressed page');
    view.setUint32(start + 20, 1, true);
    const worker = (book as any).worker as Worker;
    let cancellations = 0;
    worker.addEventListener('message', ({ data }) => {
      if (data.decompressionCancelStalled) cancellations++;
    });
    worker.postMessage({ stallDecompressionCancel: true });
    const code = await Promise.race([
      book.importBinary(bad).then(
        () => 'accepted',
        (error: any) => error.code,
      ),
      new Promise<string>((resolve) => setTimeout(() => resolve('timeout'), 5000)),
    ]);
    const after = book.revision;
    const bytesAfter = (await book.storageStats()).bytes;
    const value = await sheet.range('A1').getValues();
    const undo = await book.undo();
    const undone = await sheet.range('A1').getValues();
    const redo = await book.redo();
    const redone = await sheet.range('A1').getValues();
    worker.postMessage({ stallDecompressionCancel: false });
    await window.opensheet.open(book.id);
    book = window.opensheet.getWorkbook();
    sheet = book.getSheets()[0];
    const reopened = await sheet.range('A1').getValues();
    await book.importBinary(valid);
    const imported = await book.getSheets()[0].range('A1').getValues();
    return {
      code,
      cancellations,
      revision,
      after,
      bytesBefore,
      bytesAfter,
      value,
      undo,
      undone,
      redo,
      redone,
      reopened,
      imported,
    };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    code: 'LIMIT_EXCEEDED',
    cancellations: 1,
    revision: result.revision,
    after: result.revision,
    bytesBefore: result.bytesBefore,
    bytesAfter: result.bytesBefore,
    value: [[222]],
    undo: true,
    undone: [[111]],
    redo: true,
    redone: [[222]],
    reopened: [[222]],
    imported: [['repeated '.repeat(1000)]],
  });
});
