import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('clipboard read rejects a revision changed before the Worker read starts', async ({
  page,
}) => {
  const clipboardModule = `/@fs${fileURLToPath(new URL('../../packages/renderer/src/clipboard.ts', import.meta.url))}`;
  const structuredModule = `/@fs${fileURLToPath(new URL('../../packages/renderer/src/copy-clipboard.ts', import.meta.url))}`;
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(
    async ({ clipboardModule, structuredModule }) => {
      const book = window.opensheet.getWorkbook();
      const sheet = await book.addSheet('Copy race', { rows: 5, columns: 2 });
      await sheet.range('A1').setValues([[42]]);
      const { cutText } = await import(/* @vite-ignore */ clipboardModule);
      const { copyClipboard } = await import(/* @vite-ignore */ structuredModule);
      let resume!: () => void;
      const gate = new Promise<void>((resolve) => (resume = resolve));
      const original = book.readRange.bind(book);
      book.readRange = async (...args) => {
        await gate;
        return original(...args);
      };
      let writes = 0;
      const cut = cutText(book, sheet.range('A1'), async () => {
        writes++;
      }).then(
        () => 'success',
        (error: { code: string }) => error.code,
      );
      await sheet.insertRows(0);
      resume();
      const cutCode = await cut;
      book.readRange = original;
      let resumeStructured!: () => void;
      const structuredGate = new Promise<void>((resolve) => (resumeStructured = resolve));
      const originalCells = book.readCells.bind(book);
      book.readCells = async (...args) => {
        await structuredGate;
        return originalCells(...args);
      };
      const copy = copyClipboard(sheet.range('A2')).then(
        () => 'success',
        (error: { code: string }) => error.code,
      );
      await sheet.insertRows(0);
      resumeStructured();
      const copyCode = await copy;
      book.readCells = originalCells;
      return { cutCode, copyCode, writes, values: await sheet.range('A3').getValues() };
    },
    { clipboardModule, structuredModule },
  );
  expect(result).toEqual({
    cutCode: 'REVISION_CONFLICT',
    copyCode: 'REVISION_CONFLICT',
    writes: 0,
    values: [[42]],
  });
});

test('cut refuses a cold formula calculation limit without writing clipboard or clearing input', async ({
  page,
}) => {
  const clipboardModule = `/@fs${fileURLToPath(new URL('../../packages/renderer/src/clipboard.ts', import.meta.url))}`;
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (clipboardModule) => {
    const app = window.opensheet;
    let book = app.getWorkbook();
    const sheet = await book.addSheet('Cut limit', { rows: 130, columns: 2 });
    await sheet.range('A1').setValues([[1]]);
    await sheet
      .range('A2:A130')
      .setFormulas(Array.from({ length: 129 }, (_, i) => [`A${i + 1}+1`]));
    await sheet.range('B1').setValues([['#LIMIT!']]);
    await app.open(book.id);
    book = app.getWorkbook();
    const current = book.getSheetByName('Cut limit')!,
      revision = book.revision;
    const { cutText, copyText } = await import(/* @vite-ignore */ clipboardModule);
    let writes = 0,
      code = '';
    try {
      await cutText(book, current.range('A130'), async () => {
        writes++;
      });
    } catch (error) {
      code = (error as any).code;
    }
    return {
      code,
      writes,
      revision: book.revision,
      expectedRevision: revision,
      formula: await current.range('A130').getFormulas(),
      literal: await copyText(current.range('B1')),
    };
  }, clipboardModule);
  expect(result).toEqual({
    code: 'LIMIT_EXCEEDED',
    writes: 0,
    revision: result.expectedRevision,
    expectedRevision: result.expectedRevision,
    formula: [['A129+1']],
    literal: '#LIMIT!',
  });
});
