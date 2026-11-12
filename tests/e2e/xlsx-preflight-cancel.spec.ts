import { expect, test } from '@playwright/test';
import * as XLSX from 'xlsx';
import { fileURLToPath } from 'node:url';

test('ZIP preflight reports an oversized expansion even if decoder cancellation stalls', async ({
  page,
}) => {
  const workbook = XLSX.utils.book_new(XLSX.utils.aoa_to_sheet([['x'.repeat(1000)]]), 'Data');
  const valid = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', compression: true });
  const invalid = Buffer.from(valid);
  let directory = invalid.readUInt32LE(invalid.length - 6);
  let changed = false;
  while (directory + 46 <= invalid.length && invalid.readUInt32LE(directory) === 0x02014b50) {
    const name = invalid.readUInt16LE(directory + 28);
    const extra = invalid.readUInt16LE(directory + 30);
    const comment = invalid.readUInt16LE(directory + 32);
    if (invalid.readUInt16LE(directory + 10) === 8 && invalid.readUInt32LE(directory + 24) > 1) {
      const local = invalid.readUInt32LE(directory + 42);
      invalid.writeUInt32LE(1, directory + 24);
      invalid.writeUInt32LE(1, local + 22);
      changed = true;
      break;
    }
    directory += 46 + name + extra + comment;
  }
  expect(changed).toBe(true);
  const modulePath = `/@fs${fileURLToPath(new URL('../../apps/playground/src/zip-guard.ts', import.meta.url))}`;
  await page.route('**/zip-cancel-test.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { validateZip } from ${JSON.stringify(modulePath)};
      const Native = globalThis.DecompressionStream;
      let cancelled = false;
      globalThis.DecompressionStream = class extends Native {
        constructor(...args) {
          super(...args);
          const readable = this.readable, getReader = readable.getReader.bind(readable);
          readable.getReader = (...args) => {
            const reader = getReader(...args);
            reader.cancel = () => { cancelled = true; return new Promise(() => {}); };
            return reader;
          };
        }
      };
      self.onmessage = async ({data}) => {
        try { await validateZip(data,true); self.postMessage({ok:true,cancelled}); }
        catch (error) { self.postMessage({error:error.message,cancelled}); }
      };`,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    async ({ bad, good }) => {
      const worker = new Worker('/zip-cancel-test.worker.js', { type: 'module' });
      try {
        const send = (bytes: number[]) =>
          new Promise<any>((resolve, reject) => {
            const timer = setTimeout(
              () => reject(new Error('Worker did not report validation result')),
              5000,
            );
            worker.onmessage = ({ data }) => {
              clearTimeout(timer);
              resolve(data);
            };
            worker.onerror = (event) => {
              clearTimeout(timer);
              reject(new Error(event.message));
            };
            worker.postMessage(Uint8Array.from(bytes).buffer);
          });
        return { rejected: await send(bad), retry: await send(good) };
      } finally {
        worker.terminate();
      }
    },
    { bad: [...invalid], good: [...valid] },
  );
  expect(result.rejected).toEqual({
    error: 'Invalid ZIP: actual expanded size exceeds declared size',
    cancelled: true,
  });
  expect(result.retry).toEqual({ ok: true, cancelled: true });
});
