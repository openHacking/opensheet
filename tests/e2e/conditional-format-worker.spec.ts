import { expect, test } from '@playwright/test';

test('evaluates and cancels conditional formulas in a real Worker', async ({ page }) => {
  await page.goto('/');
  const result = await page.evaluate(async (url) => {
    const worker = new Worker(url, { type: 'module' });
    try {
      const messages: unknown[] = [];
      const next = () =>
        new Promise<any>((resolve, reject) => {
          worker.addEventListener('message', (event) => resolve(event.data), { once: true });
          worker.addEventListener('error', (event) => reject(new Error(event.message)), {
            once: true,
          });
        });
      let response = next();
      worker.postMessage('run');
      messages.push(await response);
      response = next();
      worker.postMessage('wait');
      messages.push(await response);
      response = next();
      worker.postMessage('cancel');
      messages.push(await response);
      return messages;
    } finally {
      worker.terminate();
    }
  }, `/@fs${process.cwd()}/tests/fixtures/conditional-format.worker.ts`);
  expect(result).toEqual([
    { kind: 'result', styles: [{}, { bold: true }] },
    { kind: 'waiting' },
    { kind: 'error', name: 'AbortError' },
  ]);
});
