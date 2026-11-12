import { expect, test } from '@playwright/test';

test('binary import accepts empty stream chunks and rejects trailing data atomically', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[42]]);
    const chunks: Uint8Array[] = [];
    for await (const chunk of book.streamBinary()) chunks.push(chunk);
    const bytes = new Uint8Array(await new Blob(chunks as BlobPart[]).arrayBuffer());
    const makeStream = (bad: boolean) =>
      new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(new Uint8Array());
          controller.enqueue(bytes.slice(0, 8));
          controller.enqueue(new Uint8Array());
          controller.enqueue(bytes.slice(8));
          controller.enqueue(new Uint8Array());
          if (bad) controller.enqueue(new Uint8Array([1]));
          controller.close();
        },
      });
    await sheet.range('A1').setValues([[99]]);
    await book.importBinary(makeStream(false));
    const imported = await book.getSheets()[0].range('A1').getValues();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[77]]);
    const revision = book.revision;
    let error = '';
    try {
      await book.importBinary(makeStream(true));
    } catch (e) {
      error = (e as { code: string }).code;
    }
    const after = await book.getSheets()[0].range('A1').getValues();
    return { imported, revision, afterRevision: book.revision, after, error };
  });
  expect(result.imported).toEqual([[42]]);
  expect(result.error).toBe('INVALID_ARGUMENT');
  expect(result.afterRevision).toBe(result.revision);
  expect(result.after).toEqual([[77]]);
});

test('binary import can cancel endless empty and tiny-chunk streams without publishing', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[123]]);
    const revision = book.revision;
    const errors: string[] = [];
    for (const tiny of [false, true]) {
      const abort = new AbortController();
      const header = new Uint8Array(12);
      header.set([79, 83, 66, 3]);
      new DataView(header.buffer).setUint32(4, 32 * 1024 * 1024, true);
      let first = true;
      const input = new ReadableStream<Uint8Array>({
        pull(controller) {
          controller.enqueue(tiny ? (first ? header : new Uint8Array([32])) : new Uint8Array());
          first = false;
        },
      });
      const timer = setTimeout(() => abort.abort(), 20);
      try {
        await book.importBinary(input, { signal: abort.signal });
      } catch (e) {
        errors.push((e as { code?: string; name?: string }).code ?? (e as Error).name);
      } finally {
        clearTimeout(timer);
      }
    }
    return {
      errors,
      revision,
      afterRevision: book.revision,
      value: await book.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result.errors).toEqual(['ABORTED', 'ABORTED']);
  expect(result.afterRevision).toBe(result.revision);
  expect(result.value).toEqual([[123]]);
});
