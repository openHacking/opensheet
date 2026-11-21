import { expect, test } from '@playwright/test';

test('leading U+FEFF survives page reads, history and native backups', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, openWorkbook, FileReader } = await import(/* @vite-ignore */ url);
    let source = await createWorkbook({ sheets: [{ name: 'Data', rows: 65, columns: 1 }] });
    let target = await createWorkbook({ sheets: [{ name: 'Target', rows: 1, columns: 1 }] });
    try {
      let sheet = source.getSheets()[0];
      await sheet.range('A1:A3').setValues([['\uFEFFfirst'], ['\uFEFF'], ['\uFEFFfirst']]);
      await sheet.range('A65').setValues([['\uFEFFlast']]);
      const read = async () => [
        ...(await sheet.range('A1:A3').getValues()),
        ...(await sheet.range('A65').getValues()),
      ];
      const original = await read();
      const undo = await source.undo();
      const afterUndo = await read();
      const redo = await source.redo();
      const afterRedo = await read();
      const file = await source.toJSON();
      const reader = new FileReader(file);
      const fromJSON = [
        ...reader.getSheetById(sheet.id).range('A1:A3').getValues(),
        ...reader.getSheetById(sheet.id).range('A65').getValues(),
      ];
      reader.dispose();
      const chunks: Uint8Array[] = [];
      for await (const part of source.streamBinary()) chunks.push(part);
      await target.importBinary(new Blob(chunks as BlobPart[]));
      const targetId = target.id;
      await target.close();
      target = await openWorkbook(targetId);
      const restored = target.getSheets()[0];
      const fromBinary = [
        ...(await restored.range('A1:A3').getValues()),
        ...(await restored.range('A65').getValues()),
      ];
      const sourceId = source.id;
      await source.close();
      source = await openWorkbook(sourceId);
      sheet = source.getSheets()[0];
      return {
        original,
        undo,
        afterUndo,
        redo,
        afterRedo,
        fromJSON,
        fromBinary,
        reopened: await read(),
      };
    } finally {
      await source.deleteStorage();
      await target.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  const values = [['\uFEFFfirst'], ['\uFEFF'], ['\uFEFFfirst'], ['\uFEFFlast']];
  expect(result).toEqual({
    original: values,
    undo: true,
    afterUndo: [...values.slice(0, 3), [null]],
    redo: true,
    afterRedo: values,
    fromJSON: values,
    fromBinary: values,
    reopened: values,
  });
});
