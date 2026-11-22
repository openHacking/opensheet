import { expect, test } from '@playwright/test';

test('rejects string arenas with an ignored prefix before replacing a workbook', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (root) => {
    const { writeBinary } = await import(/* @vite-ignore */ `${root}/binary-file.ts`);
    const { verifyBinary } = await import(/* @vite-ignore */ `${root}/binary-verify.ts`);
    const { encodeBlock, decodeBlock } = await import(/* @vite-ignore */ `${root}/block-codec.ts`);
    const book = window.opensheet.getWorkbook(),
      sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[111]]);
    await sheet.range('A1').setValues([[222]]);
    const revision = book.revision,
      metadata = await book.toJSON();
    metadata.sheets.forEach((entry: any) => (entry.blocks = []));
    const cases: Array<{ type: number; errors: string[] }> = [];
    let validFile: Uint8Array | undefined;
    for (const wide of [false, true]) {
      for (const dictionary of [false, true]) {
        const text = wide ? '\ud800abcdef' : 'abcdef';
        const cells: Record<string, any> = {};
        for (let row = 0; row < (dictionary ? 12 : 1); row++)
          cells[`${row}:0`] = { rowId: row, columnId: 0, input: { type: 'string', value: text } };
        const encoded = encodeBlock(cells, [], [], {}),
          view = new DataView(encoded.buffer);
        const headerLength = view.getUint32(4, true),
          base = Math.ceil((8 + headerLength) / 8) * 8;
        const meta = JSON.parse(new TextDecoder().decode(encoded.subarray(8, 8 + headerLength)));
        const section = meta.sections.find((entry: any) => entry.name === 'strings');
        const directory =
          base + section.offset + 4 + (dictionary ? section.positions.count * 2 : 0);
        if (section.type !== (wide ? 16 : 0) + (dictionary ? 1 : 0))
          throw new Error('Wrong encoding');
        const build = async (bytes: Uint8Array) => {
          async function* pages() {
            yield { sheet: 0, row: 0, column: 0, bytes };
          }
          const parts: BlobPart[] = [];
          for await (const part of writeBinary(metadata, pages())) parts.push(part);
          return new Uint8Array(await new Blob(parts).arrayBuffer());
        };
        const valid = await build(encoded);
        const verified = await verifyBinary(valid);
        if (verified.cells !== (dictionary ? 12 : 1)) throw new Error('Valid encoding rejected');
        if (!wide && !dictionary) validFile = valid;
        view.setUint32(directory, wide ? 2 : 1, true);
        const invalid = await build(encoded);
        const errors: string[] = [];
        for (const run of [
          () => decodeBlock(encoded),
          () => verifyBinary(invalid),
          () => book.importBinary(invalid),
        ]) {
          try {
            await run();
            errors.push('accepted');
          } catch (error) {
            errors.push((error as any).code);
          }
        }
        cases.push({ type: section.type, errors });
      }
    }
    const after = book.revision,
      value = await sheet.range('A1').getValues();
    const undo = await book.undo(),
      undone = await sheet.range('A1').getValues();
    const redo = await book.redo(),
      redone = await sheet.range('A1').getValues();
    await window.opensheet.open(book.id);
    const reopened = window.opensheet.getWorkbook(),
      reopenedValue = await reopened.getSheets()[0].range('A1').getValues();
    await reopened.importBinary(validFile!);
    return {
      cases,
      revision,
      after,
      value,
      undo,
      undone,
      redo,
      redone,
      reopenedValue,
      importedValue: await reopened.getSheets()[0].range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src`);
  expect(result.cases).toEqual(
    [0, 1, 16, 17].map((type) => ({ type, errors: Array(3).fill('CORRUPT_STORAGE') })),
  );
  expect(result.after).toBe(result.revision);
  expect(result.value).toEqual([[222]]);
  expect(result.undo).toBe(true);
  expect(result.undone).toEqual([[111]]);
  expect(result.redo).toBe(true);
  expect(result.redone).toEqual([[222]]);
  expect(result.reopenedValue).toEqual([[222]]);
  expect(result.importedValue).toEqual([['abcdef']]);
});
