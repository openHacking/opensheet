import { expect, test } from '@playwright/test';

for (const failure of ['error', 'messageerror'] as const) {
  test(`failed Worker releases its writer lock after ${failure}`, async ({ page, context }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    await page.evaluate(async () => {
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1')
        .setValues([[765]]);
    });
    const reader = await context.newPage();
    await reader.goto('/#budget');
    await expect(reader.locator('#take-write-lock')).toBeVisible();
    await page.evaluate((kind) => {
      const worker = (window.opensheet.getWorkbook() as any).worker as Worker;
      if (kind === 'error')
        worker.onerror?.call(worker, new ErrorEvent('error', { message: 'Injected worker fault' }));
      else worker.onmessageerror?.call(worker, new MessageEvent('messageerror'));
    }, failure);
    await expect(page.locator('#save-state')).toContainText('Reopen workbook');
    await expect
      .poll(() => reader.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()), {
        timeout: 10000,
      })
      .toBe(true);
    expect(
      await reader.evaluate(async () =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[765]]);
    await reader.evaluate(async () => {
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1')
        .setValues([[766]]);
      await window.opensheet.open(window.opensheet.getWorkbook().id);
    });
    expect(
      await reader.evaluate(async () =>
        window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues(),
      ),
    ).toEqual([[766]]);
    await reader.close();
  });
}

test('a read-only tab can take over after the writer closes, reloading the latest head', async ({
  page,
  context,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const second = await context.newPage();
  await second.goto('/#budget');
  await expect(second.locator('#take-write-lock')).toBeVisible();
  await second.locator('#take-write-lock').click();
  await expect(second.locator('#notice')).toContainText('Another tab is still editing');
  await expect(second.locator('.os-grid')).toHaveAttribute('aria-readonly', 'true');
  await second.evaluate(() => {
    ((window.opensheet.getWorkbook() as any).channel as BroadcastChannel).onmessage = null;
  });
  await page.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[1234]]);
  });
  await page.close();
  await expect
    .poll(() => second.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()), {
      timeout: 10000,
    })
    .toBe(true);
  await expect(second.locator('#take-write-lock')).toBeHidden();
  await expect(second.locator('.os-grid')).toHaveAttribute('aria-readonly', 'false');
  expect(
    await second.evaluate(
      async () =>
        (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
    ),
  ).toBe(1234);
  await second.evaluate(async () => {
    await window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[5678]]);
  });
  const third = await context.newPage();
  await third.goto('/#budget');
  await expect(third.locator('#take-write-lock')).toBeVisible();
  expect(
    await third.evaluate(
      async () =>
        (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
    ),
  ).toBe(5678);
  await third.locator('#take-write-lock').click();
  await expect(third.locator('#notice')).toContainText('Another tab is still editing');
  await third.close();
  await second.close();
});

test('competing readers cannot both become writers', async ({ page, context }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const readers = await Promise.all([context.newPage(), context.newPage()]);
  await Promise.all(readers.map((reader) => reader.goto('/#budget')));
  await Promise.all(
    readers.map((reader) => expect(reader.locator('#take-write-lock')).toBeVisible()),
  );
  await page.evaluate(() => window.opensheet.getWorkbook().close());
  const attempts = await Promise.all(
    readers.map((reader) =>
      reader.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()),
    ),
  );
  expect(attempts.filter(Boolean)).toHaveLength(1);
  const [winner, loser] = attempts[0] ? readers : [readers[1], readers[0]];
  await winner.evaluate(async () =>
    window.opensheet
      .getWorkbook()
      .getSheets()[0]
      .range('A1')
      .setValues([[42]]),
  );
  const error = await loser.evaluate(async () => {
    try {
      await window.opensheet
        .getWorkbook()
        .getSheets()[0]
        .range('A1')
        .setValues([[99]]);
      return 'wrote';
    } catch (e) {
      return (e as { code?: string }).code;
    }
  });
  expect(error).toBe('READ_ONLY');
  await winner.close();
  await expect
    .poll(() => loser.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()))
    .toBe(true);
  expect(
    await loser.evaluate(
      async () =>
        (await window.opensheet.getWorkbook().getSheets()[0].range('A1').getValues())[0][0],
    ),
  ).toBe(42);
  await loser.close();
});

test('failed recovery releases the lock for another reader', async ({ page, context }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const readers = await Promise.all([context.newPage(), context.newPage()]);
  await Promise.all(readers.map((reader) => reader.goto('/#budget')));
  await Promise.all(
    readers.map((reader) => expect(reader.locator('#take-write-lock')).toBeVisible()),
  );
  await page.evaluate(() => window.opensheet.getWorkbook().close());
  await readers[0].evaluate(async () => {
    const id = window.opensheet.getWorkbook().id;
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const read = db.transaction('heads').objectStore('heads').get(id);
    (window as any).savedHead = await new Promise((resolve, reject) => {
      read.onsuccess = () => resolve(read.result);
      read.onerror = () => reject(read.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').delete(id);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
  });
  expect(
    await readers[0].evaluate(async () => {
      try {
        await window.opensheet.getWorkbook().tryAcquireWriteLock();
        return 'unexpected success';
      } catch (error) {
        return (error as { code?: string }).code;
      }
    }),
  ).toBe('CORRUPT_STORAGE');
  expect(await readers[0].evaluate(() => window.opensheet.getWorkbook().isReadOnly)).toBe(true);
  await readers[0].evaluate(async () => {
    const id = window.opensheet.getWorkbook().id;
    const open = indexedDB.open('opensheet');
    const db = await new Promise<IDBDatabase>((resolve, reject) => {
      open.onsuccess = () => resolve(open.result);
      open.onerror = () => reject(open.error);
    });
    const tx = db.transaction('heads', 'readwrite');
    tx.objectStore('heads').put((window as any).savedHead, id);
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error);
    });
    db.close();
  });
  await expect
    .poll(() => readers[1].evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()))
    .toBe(true);
  await Promise.all(readers.map((reader) => reader.close()));
});

test('explicit read-only mode never exposes the handover action', async ({ page, context }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const second = await context.newPage();
  await second.goto('/#budget');
  await expect(second.locator('#scene-title')).toHaveText('Budget Simulator');
  await expect(second.locator('#take-write-lock')).toBeVisible();
  await second.evaluate(() => window.opensheet.setMode('read'));
  await expect(second.locator('#take-write-lock')).toBeHidden();
  await page.evaluate(() => window.opensheet.getWorkbook().close());
  expect(
    await second.evaluate(async () => {
      try {
        await window.opensheet.getWorkbook().tryAcquireWriteLock();
        return 'acquired';
      } catch (error) {
        return (error as { code?: string }).code;
      }
    }),
  ).toBe('READ_ONLY');
  await expect(second.locator('.os-grid')).toHaveAttribute('aria-readonly', 'true');
  await second.evaluate(() => window.opensheet.setMode('edit'));
  await expect(second.locator('#take-write-lock')).toBeVisible();
  await expect
    .poll(() => second.evaluate(() => window.opensheet.getWorkbook().tryAcquireWriteLock()))
    .toBe(true);
  await second.evaluate(() => window.opensheet.setMode('read'));
  expect(
    await second.evaluate(async () => {
      try {
        await window.opensheet.getWorkbook().tryAcquireWriteLock();
        return 'acquired';
      } catch (error) {
        return (error as { code?: string }).code;
      }
    }),
  ).toBe('READ_ONLY');
  await expect(second.locator('.os-grid')).toHaveAttribute('aria-readonly', 'true');
  await second.close();
});
