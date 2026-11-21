import { expect, test } from '@playwright/test';

for (const format of ['json', 'binary'] as const) {
  test(`${format} import reports invalid transferred input when cancellation stalls`, async ({
    page,
  }) => {
    test.setTimeout(30_000);
    await page.goto('/');
    const result = await page.evaluate(
      async ({ url, format }) => {
        const { createWorkbook, openWorkbook } = await import(/* @vite-ignore */ url);
        let book = await createWorkbook({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
        try {
          let sheet = book.getSheets()[0];
          await sheet.range('A1').setValues([[111]]);
          await sheet.range('A1').setValues([[222]]);
          const revision = book.revision;
          const bytes = (await book.storageStats()).bytes;
          let count = 0;
          let started!: () => void;
          const cancelled = new Promise<void>((resolve) => {
            started = resolve;
          });
          const input = new ReadableStream<Uint8Array>({
            start(controller) {
              controller.enqueue(format === 'json' ? new Uint8Array([255]) : new Uint8Array(12));
            },
            cancel() {
              count++;
              started();
              return new Promise<void>(() => {});
            },
          });
          const operation = format === 'json' ? book.importJSON(input) : book.importBinary(input);
          const code = await Promise.race([
            operation.then(
              () => 'accepted',
              (error: { code?: string }) => error.code,
            ),
            new Promise<string>((resolve) => setTimeout(() => resolve('timeout'), 5000)),
          ]);
          if (code === 'timeout') throw new Error(`${format} import failed to settle`);
          await Promise.race([
            cancelled,
            new Promise<never>((_, reject) =>
              setTimeout(() => reject(new Error('Cancel not requested')), 3000),
            ),
          ]);
          const after = book.revision;
          const afterBytes = (await book.storageStats()).bytes;
          const value = await sheet.range('A1').getValues();
          const undo = await book.undo();
          const undone = await sheet.range('A1').getValues();
          const redo = await book.redo();
          const redone = await sheet.range('A1').getValues();
          const id = book.id;
          await book.close();
          book = await openWorkbook(id);
          sheet = book.getSheets()[0];
          const reopened = await sheet.range('A1').getValues();
          await sheet.range('A1').setValues([[333]]);
          return {
            code,
            count,
            revision,
            after,
            bytes,
            afterBytes,
            value,
            undo,
            undone,
            redo,
            redone,
            reopened,
            retry: await sheet.range('A1').getValues(),
          };
        } finally {
          await book.deleteStorage();
        }
      },
      { url: `/@fs${process.cwd()}/packages/core/src/index.ts`, format },
    );
    expect(result).toEqual({
      code: 'INVALID_ARGUMENT',
      count: 1,
      revision: result.revision,
      after: result.revision,
      bytes: result.bytes,
      afterBytes: result.bytes,
      value: [[222]],
      undo: true,
      undone: [[111]],
      redo: true,
      redone: [[222]],
      reopened: [[222]],
      retry: [[333]],
    });
  });
}

test('backup comparison reports mismatch while input cancellation is stalled', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const { createWorkbook, compareBinaryBackup } = await import(/* @vite-ignore */ url);
    const book = await createWorkbook({ sheets: [{ name: 'Data', rows: 2, columns: 2 }] });
    try {
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[222]]);
      const parts: Uint8Array[] = [];
      for await (const part of book.streamBinary()) parts.push(part);
      const blob = new Blob(parts as BlobPart[]);
      const original = blob.stream.bind(blob);
      let calls = 0;
      let cancelled = 0;
      const mismatch = new ReadableStream<Uint8Array<ArrayBuffer>>({
        start(controller) {
          controller.enqueue(new Uint8Array([255]));
        },
        cancel() {
          cancelled++;
          return new Promise<void>(() => {});
        },
      });
      Object.defineProperty(blob, 'stream', {
        value: () => (++calls === 1 ? original() : mismatch),
      });
      const comparison = await Promise.race([
        compareBinaryBackup(book, blob),
        new Promise<'timeout'>((resolve) => setTimeout(() => resolve('timeout'), 5000)),
      ]);
      return {
        comparison,
        revision: book.revision,
        cancelled,
        locked: mismatch.locked,
        value: await book.getSheets()[0].range('A1').getValues(),
      };
    } finally {
      await book.deleteStorage();
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    comparison: { matches: false, revision: result.revision },
    revision: result.revision,
    cancelled: 1,
    locked: false,
    value: [[222]],
  });
});
