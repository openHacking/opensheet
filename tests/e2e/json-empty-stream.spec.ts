import { expect, test } from '@playwright/test';

test('JSON import rejects invalid UTF-8 and cancels nonproductive streams atomically', async ({
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
    const invalid = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(new Uint8Array([0xff]));
        controller.close();
      },
    });
    try {
      await book.importJSON(invalid);
    } catch (e) {
      errors.push((e as { code: string }).code);
    }
    for (const chunk of [new Uint8Array(), new TextEncoder().encode('   ')]) {
      const controller = new AbortController();
      const input = new ReadableStream<Uint8Array>({
        pull(stream) {
          stream.enqueue(chunk);
        },
      });
      const timer = setTimeout(() => controller.abort(), 20);
      try {
        await book.importJSON(input, { signal: controller.signal });
      } catch (e) {
        errors.push((e as { code: string }).code);
      } finally {
        clearTimeout(timer);
      }
    }
    {
      const controller = new AbortController();
      let first = true;
      const encoder = new TextEncoder();
      const input = new ReadableStream<Uint8Array>({
        pull(stream) {
          stream.enqueue(encoder.encode(first ? '[0,' : '0,'));
          first = false;
        },
      });
      const timer = setTimeout(() => controller.abort(), 20);
      try {
        await book.importJSON(input, { signal: controller.signal });
      } catch (e) {
        errors.push((e as { code: string }).code);
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
  expect(result.errors).toEqual(['INVALID_ARGUMENT', 'ABORTED', 'ABORTED', 'ABORTED']);
  expect(result.afterRevision).toBe(result.revision);
  expect(result.value).toEqual([[123]]);
});

test('JSON import decodes large chunks across UTF-8 boundaries and rejects oversized tokens atomically', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['😀']]);
    const text = JSON.stringify(await book.toJSON());
    const index = text.indexOf('😀');
    if (index < 0) throw new Error('Missing Unicode fixture');
    const encoder = new TextEncoder();
    const before = encoder.encode(text.slice(0, index)).length;
    const padding = (65535 - (before % 65536) + 65536) % 65536;
    const bytes = encoder.encode(' '.repeat(padding) + text);
    const source = (data: Uint8Array) =>
      new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(data);
          controller.close();
        },
      });
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([['other']]);
    await book.importJSON(source(bytes));
    const imported = await book.getSheets()[0].range('A1').getValues();
    const revision = book.revision;
    const invalid = encoder.encode('{"value":"' + 'x'.repeat(2 * 1024 * 1024) + '"}');
    let error = '';
    try {
      await book.importJSON(source(invalid));
    } catch (e) {
      error = (e as { code: string }).code;
    }
    return {
      imported,
      error,
      revision,
      afterRevision: book.revision,
      after: await book.getSheets()[0].range('A1').getValues(),
    };
  });
  expect(result.imported).toEqual([['😀']]);
  expect(result.error).toBe('LIMIT_EXCEEDED');
  expect(result.afterRevision).toBe(result.revision);
  expect(result.after).toEqual([['😀']]);
});
