import { expect, test } from '@playwright/test';

test('generation is exclusive and transactions cannot start behind a pending write', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const snapshot = await book.toJSON();
    const worker = (book as any).worker as Worker;
    const send = worker.postMessage.bind(worker);
    let held: any;
    let holdMethod = 'execute';
    worker.postMessage = ((message: any) => {
      if (message.method === holdMethod) held = message;
      else send(message);
    }) as typeof worker.postMessage;
    const code = async (operation: () => Promise<unknown>) => {
      try {
        await operation();
        return 'unexpected success';
      } catch (error) {
        return (error as any).code;
      }
    };
    const write = sheet.range('A1').setValues([[55]]);
    const pendingTransaction = await code(() => book.transaction({}, () => {}));
    send(held);
    await write;
    holdMethod = 'generate';
    const generation = book.generate(100);
    const blocked = [];
    blocked.push(await code(() => sheet.range('A1').setValues([[99]])));
    blocked.push(await code(() => book.transaction({}, () => {})));
    blocked.push(await code(() => book.undo()));
    blocked.push(await code(() => book.redo()));
    blocked.push(await code(() => book.importJSON(snapshot)));
    blocked.push(await code(() => book.deleteStorage()));
    blocked.push(await code(() => book.generate(1)));
    send(held);
    await generation;
    worker.postMessage = send as typeof worker.postMessage;
    return {
      pendingTransaction,
      blocked,
      generated: (await book.getSheets()[0].range('A1').getValues())[0][0],
    };
  });
  expect(result).toEqual({
    pendingTransaction: 'BUSY',
    blocked: Array(7).fill('BUSY'),
    generated: 0,
  });
});

test('transaction-scoped subscriptions are removed on commit and rollback', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    let calls = 0;
    const shared = () => {
      calls++;
    };
    const stopParent = book.onCommit(shared);
    const counts = () =>
      ['listeners', 'dataListeners', 'progressListeners', 'persistenceListeners'].map(
        (name) => (book as any)[name].size,
      );
    const baseline = counts();
    let stale: typeof book | undefined;
    const attach = (scope: typeof book) => {
      stale = scope;
      scope.onCommit(shared);
      scope.onData(() => {});
      scope.onProgress(() => {});
      scope.onPersistence(() => {});
    };
    await book.transaction({}, async (scope) => {
      attach(scope);
      await scope
        .getSheets()[0]
        .range('A1')
        .setValues([[12]]);
    });
    const afterCommit = counts();
    const commitCalls = calls;
    const expired: string[] = [];
    for (const register of [
      () => stale!.onCommit(shared),
      () => stale!.onData(() => {}),
      () => stale!.onProgress(() => {}),
      () => stale!.onPersistence(() => {}),
    ]) {
      try {
        register();
        expired.push('unexpected success');
      } catch (error) {
        expired.push((error as any).code);
      }
    }
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[13]]);
    const laterCalls = calls;
    let rollback = '';
    try {
      await book.transaction({}, (scope) => {
        attach(scope);
        throw new Error('rollback');
      });
    } catch (error) {
      rollback = (error as Error).message;
    }
    const afterRollback = counts();
    stopParent();
    return { baseline, afterCommit, commitCalls, expired, laterCalls, rollback, afterRollback };
  });
  expect(result.afterCommit).toEqual(result.baseline);
  expect(result.afterRollback).toEqual(result.baseline);
  expect(result.commitCalls).toBe(2);
  expect(result.laterCalls).toBe(3);
  expect(result.expired).toEqual(Array(4).fill('DISPOSED'));
  expect(result.rollback).toBe('rollback');
});

test('expired transaction scopes cannot issue or complete Worker reads', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const book = window.opensheet.getWorkbook();
    await book
      .getSheets()[0]
      .range('A1')
      .setValues([[4321]]);
    const sheetId = book.getSheets()[0].id;
    const worker = (book as any).worker as Worker;
    const send = worker.postMessage.bind(worker);
    let held: any;
    let reads = 0;
    worker.postMessage = ((message: any) => {
      if (message.method === 'search') {
        reads++;
        held = message;
      } else send(message);
    }) as typeof worker.postMessage;
    let stale: typeof book | undefined;
    let pending!: Promise<string>;
    let readyIsScope = false;
    await book.transaction({}, async (scope) => {
      stale = scope;
      readyIsScope = (await scope.ready()) === scope;
      pending = scope.search(sheetId, 'budget').then(
        () => 'unexpected success',
        (error) => error.code,
      );
    });
    send(held);
    const late = await pending;
    const rejected: string[] = [];
    for (const action of [
      () => stale!.search(sheetId, 'budget'),
      () => stale!.scanCells(sheetId),
      () => stale!.storageStats(),
      () => stale!.ready(),
      () => stale!.calculation.calculate(),
      () => stale!.calculation.getState(),
    ]) {
      try {
        await action();
        rejected.push('unexpected success');
      } catch (error) {
        rejected.push((error as any).code);
      }
    }
    worker.postMessage = send as typeof worker.postMessage;
    const parentValue = await book.getSheets()[0].range('A1').getValues();
    return { readyIsScope, late, rejected, reads, parentValue };
  });
  expect(result).toEqual({
    readyIsScope: true,
    late: 'DISPOSED',
    rejected: Array(6).fill('DISPOSED'),
    reads: 1,
    parentValue: [[4321]],
  });
});

test('scoped cancellation aborts its transaction and an expired scope cannot cancel another', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = book.getSheets()[0];
    await sheet.range('A1').setValues([[10]]);
    const revision = book.revision;
    let oldScope: typeof book | undefined;
    let abortCode = '';
    try {
      await book.transaction({}, async (scope) => {
        oldScope = scope;
        await scope
          .getSheets()[0]
          .range('A1')
          .setValues([[20]]);
        scope.cancel();
      });
    } catch (error) {
      abortCode = (error as any).code;
    }
    const afterCancel = { revision: book.revision, value: await sheet.range('A1').getValues() };
    let expiredCode = '';
    await book.transaction({}, async (scope) => {
      await scope
        .getSheets()[0]
        .range('A1')
        .setValues([[30]]);
      try {
        oldScope!.cancel();
      } catch (error) {
        expiredCode = (error as any).code;
      }
    });
    const afterNew = await sheet.range('A1').getValues();
    await app.open(book.id);
    const reopened = await app.getWorkbook().getSheets()[0].range('A1').getValues();
    return { revision, abortCode, afterCancel, expiredCode, afterNew, reopened };
  });
  expect(result).toEqual({
    revision: result.revision,
    abortCode: 'ABORTED',
    afterCancel: { revision: result.revision, value: [[10]] },
    expiredCode: 'DISPOSED',
    afterNew: [[30]],
    reopened: [[30]],
  });
});
