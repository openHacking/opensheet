import { expect, test } from '@playwright/test';

test('imports and reopens valid inline native page values', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/index.ts`);
    const { encodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[222]]);
    const metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const bytes = encodeBlock(
      { '0:0': { rowId: 0, columnId: 0, input: { type: 'number', value: 4 } } },
      [],
      [],
      {},
    );
    const length = new DataView(bytes.buffer, bytes.byteOffset).getUint32(4, true);
    const meta = JSON.parse(new TextDecoder().decode(bytes.subarray(8, 8 + length)));
    meta.sections = [];
    meta.data.numbers = { positions: [0], values: [4] };
    meta.data.strings = { positions: [1], values: ['hello'] };
    meta.data.booleans = { positions: [2], values: [false] };
    meta.data.formulas = { positions: [3], values: ['A1+1'] };
    meta.data.errors = { positions: [4], values: ['#REF!'] };
    const header = new TextEncoder().encode(JSON.stringify(meta));
    const inline = new Uint8Array(Math.ceil((8 + header.length) / 8) * 8);
    inline.set(bytes.subarray(0, 4));
    new DataView(inline.buffer).setUint32(4, header.length, true);
    inline.set(header, 8);
    async function* pages() {
      yield { sheet: 0, row: 0, column: 0, bytes: inline };
    }
    const parts: BlobPart[] = [];
    for await (const part of writeBinary(metadata, pages())) parts.push(part);
    const file = new Blob(parts);
    const verified = await verifyBinary(file);
    await book.importBinary(file);
    const values = await book.getSheets()[0].range('A1:E1').getValues();
    const formula = await book.getSheets()[0].range('D1').getFormulas();
    const clearedHistory = await book.undo();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[8]]);
    await book.undo();
    const undone = await book.getSheets()[0].range('A1').getValues();
    await book.redo();
    await app.open(book.id);
    const reopened = await app.getWorkbook().getSheets()[0].range('A1:E1').getValues();
    return { verified: verified.cells, values, formula, clearedHistory, undone, reopened };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result).toEqual({
    verified: 5,
    values: [[4, 'hello', false, 5, { error: '#REF!' }]],
    formula: [['A1+1']],
    clearedHistory: false,
    undone: [[4]],
    reopened: [[8, 'hello', false, 9, { error: '#REF!' }]],
  });
});
