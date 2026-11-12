import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

test('cancels an internal staged validation waiting on a Worker read', async ({ page }) => {
  const modulePath = fileURLToPath(
    new URL('../../packages/core/src/validation-batch.ts', import.meta.url),
  );
  await page.route('**/staged-validation-cancel.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { checkStagedValidation } from ${JSON.stringify(`/@fs${modulePath}`)};
      const controller = new AbortController();
      const rule = { kind: 'number', operator: 'gte', first: 1, allowBlank: false, policy: 'reject' };
      const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
      const checking = checkStagedValidation([{id: 'required', range, rule}], [{row: 0, column: 0}], () => {
        self.postMessage('reading');
        return new Promise(() => {});
      }, controller.signal);
      self.onmessage = () => controller.abort();
      checking.then(() => self.postMessage('unexpected success'), error => self.postMessage(error.name));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise<string[]>((resolve, reject) => {
        const worker = new Worker('/staged-validation-cancel.worker.js', { type: 'module' });
        const messages: string[] = [];
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker did not cancel'));
        }, 5000);
        worker.onerror = () => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error('Worker failed'));
        };
        worker.onmessage = ({ data }) => {
          messages.push(data);
          if (data === 'reading') worker.postMessage('cancel');
          else {
            clearTimeout(timer);
            worker.terminate();
            resolve(messages);
          }
        };
      }),
  );
  expect(result).toEqual(['reading', 'AbortError']);
});

test('internal validation rejects a calculated error even when its rule is true', async ({
  page,
}) => {
  const batchPath = fileURLToPath(
    new URL('../../packages/core/src/validation-batch.ts', import.meta.url),
  );
  const calcPath = fileURLToPath(
    new URL('../../packages/core/src/async-calculation.ts', import.meta.url),
  );
  await page.route('**/staged-validation-error.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { checkStagedValidation } from ${JSON.stringify(`/@fs${batchPath}`)};
      import { AsyncCalculation } from ${JSON.stringify(`/@fs${calcPath}`)};
      const calc = new AsyncCalculation(async (_sheet, row, column) => ({
        rowId: row, columnId: column, input: { type: 'formula', expression: '1/0' }
      }), () => undefined);
      const session = calc.validationSession().forSheet('data');
      const range = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 };
      const rule = { kind: 'formula', expression: 'TRUE', allowBlank: false, policy: 'reject' };
      checkStagedValidation([{ id: 'f', range, rule }], [{ row: 0, column: 0 }], session.read, undefined, session.evaluateFormula)
        .then(result => self.postMessage({ result }), error => self.postMessage({ error: String(error) }));
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise((resolve, reject) => {
        const worker = new Worker('/staged-validation-error.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker did not finish'));
        }, 5000);
        worker.onerror = () => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error('Worker failed'));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result).toEqual({
    result: {
      rejected: 1,
      warnings: 0,
      details: [{ row: 0, column: 0, ruleId: 'f', policy: 'reject' }],
      truncated: false,
    },
  });
});

test('cancels an uncovered validation scan in a real Worker', async ({ page }) => {
  const modulePath = fileURLToPath(
    new URL('../../packages/core/src/validation-batch.ts', import.meta.url),
  );
  await page.route('**/validation-scan-cancel.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { checkStagedValidation } from ${JSON.stringify(`/@fs${modulePath}`)};
      const controller = new AbortController();
      const rule = { kind: 'number', operator: 'gte', first: 1, allowBlank: false, policy: 'reject' };
      const range = { startRow: 1000, endRow: 1001, startColumn: 0, endColumn: 1 };
      const rules = Array.from({ length: 900 }, (_, i) => ({ id: String(i), range, rule }));
      const positions = Array.from({ length: 1000 }, (_, row) => ({ row, column: 0 }));
      let reads = 0;
      const read = async () => { reads++; return 1; };
      let clock = 0;
      Object.defineProperty(performance, 'now', { configurable: true, value: () => clock += 10 });
      setTimeout(() => controller.abort(), 0);
      try {
        await checkStagedValidation(rules, positions, read, controller.signal);
        self.postMessage({ unexpected: true });
      } catch (error) {
        delete performance.now;
        const result = await checkStagedValidation(rules, positions, read);
        self.postMessage({ name: error.name, reads, result });
      }
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise((resolve, reject) => {
        const worker = new Worker('/validation-scan-cancel.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker did not finish'));
        }, 5000);
        worker.onerror = () => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error('Worker failed'));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result).toEqual({
    name: 'AbortError',
    reads: 0,
    result: { rejected: 0, warnings: 0, details: [], truncated: false },
  });
});

test('cancels while validating positions before a later invalid position', async ({ page }) => {
  const modulePath = fileURLToPath(
    new URL('../../packages/core/src/validation-batch.ts', import.meta.url),
  );
  await page.route('**/validation-position-cancel.worker.js', (route) =>
    route.fulfill({
      contentType: 'text/javascript',
      body: `
      import { checkStagedValidation } from ${JSON.stringify(`/@fs${modulePath}`)};
      const controller = new AbortController();
      const positions = Array.from({length: 65536}, (_, row) => ({row, column: 0}));
      positions[positions.length - 1].column = -1;
      let reads = 0, clock = 0;
      const read = async () => { reads++; return 1; };
      Object.defineProperty(performance, 'now', {configurable:true, value: () => clock += 10});
      setTimeout(() => controller.abort(), 0);
      try {
        await checkStagedValidation([], positions, read, controller.signal);
        self.postMessage({unexpected:true});
      } catch (error) {
        delete performance.now;
        let retry;
        try { await checkStagedValidation([], positions, read); retry = 'success'; }
        catch (failure) { retry = failure.code; }
        positions[positions.length - 1].column = 0;
        const result = await checkStagedValidation([], positions, read);
        self.postMessage({name:error.name, retry, reads, result});
      }
    `,
    }),
  );
  await page.goto('/');
  const result = await page.evaluate(
    () =>
      new Promise((resolve, reject) => {
        const worker = new Worker('/validation-position-cancel.worker.js', { type: 'module' });
        const timer = setTimeout(() => {
          worker.terminate();
          reject(new Error('Worker timed out'));
        }, 10000);
        worker.onerror = () => {
          clearTimeout(timer);
          worker.terminate();
          reject(new Error('Worker failed'));
        };
        worker.onmessage = ({ data }) => {
          clearTimeout(timer);
          worker.terminate();
          resolve(data);
        };
      }),
  );
  expect(result).toEqual({
    name: 'AbortError',
    retry: 'INVALID_ARGUMENT',
    reads: 0,
    result: { rejected: 0, warnings: 0, details: [], truncated: false },
  });
});
