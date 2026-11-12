import { expect, test } from '@playwright/test';

test('compares a saved backup to a live revision without changing data', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (moduleUrl) => {
    const { compareBinaryBackup } = await import(/* @vite-ignore */ moduleUrl);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[101]]);
    const revision = book.revision;
    const parts: BlobPart[] = [];
    for await (const part of book.streamBinary()) parts.push(part as BlobPart);
    const backup = new Blob(parts);
    const equal = await compareBinaryBackup(book, backup);
    const afterEqual = book.revision;
    await sheet.range('A1').setValues([[202]]);
    const different = await compareBinaryBackup(book, backup);
    const controller = new AbortController();
    controller.abort();
    let aborted = '';
    try {
      await compareBinaryBackup(book, backup, { signal: controller.signal });
    } catch (error) {
      aborted = (error as Error).name;
    }
    const invalid = new Blob([new Uint8Array([1, 2, 3])]);
    let invalidRejected = false;
    try {
      await compareBinaryBackup(book, invalid);
    } catch {
      invalidRejected = true;
    }
    return {
      revision,
      equal,
      afterEqual,
      different,
      aborted,
      invalidRejected,
      value: await sheet.range('A1').getValues(),
    };
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({
    revision: result.revision,
    equal: { matches: true, revision: result.revision },
    afterEqual: result.revision,
    different: { matches: false, revision: result.revision + 1 },
    aborted: 'AbortError',
    invalidRejected: true,
    value: [[202]],
  });
});

test('comparison rejects a workbook edited during export', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (moduleUrl) => {
    const { compareBinaryBackup } = await import(/* @vite-ignore */ moduleUrl);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const chunks: BlobPart[] = [];
    for await (const chunk of book.streamBinary()) chunks.push(chunk as BlobPart);
    const backup = new Blob(chunks);
    const stream = book.streamBinary.bind(book);
    book.streamBinary = async function* (options = {}) {
      let first = true;
      for await (const chunk of stream(options)) {
        yield chunk;
        if (first) {
          first = false;
          await sheet.range('A1').setValues([[777]]);
        }
      }
    };
    try {
      await compareBinaryBackup(book, backup);
      return { code: 'unexpected success', value: await sheet.range('A1').getValues() };
    } catch (error) {
      return { code: (error as any).code, value: await sheet.range('A1').getValues() };
    } finally {
      book.streamBinary = stream;
    }
  }, `/@fs${process.cwd()}/packages/core/src/index.ts`);
  expect(result).toEqual({ code: 'REVISION_CONFLICT', value: [[777]] });
});
