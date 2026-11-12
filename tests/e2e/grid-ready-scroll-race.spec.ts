import { expect, test } from '@playwright/test';

test('grid readiness follows a scroll arriving before the paint frame', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = await book.addSheet('Ready race', { rows: 150, columns: 2 });
    await sheet.range('A101').setValues([['target']]);
    const grid = app.getGrid()!;
    grid.setSheet(sheet.id);
    await grid.ready();
    const original = window.requestAnimationFrame;
    let moves = 0;
    window.requestAnimationFrame = (callback) =>
      original((time) => {
        if (moves < 3) grid.scroller.scrollTop = [1980, 2400, 3000][moves++];
        callback(time);
      });
    try {
      await grid.ready();
      return {
        moves,
        scrollTop: grid.scroller.scrollTop,
        loaded: book.isLoaded(sheet.id, 100, 0),
        value: book.peekValue(sheet.id, 100, 0),
      };
    } finally {
      window.requestAnimationFrame = original;
    }
  });
  expect(result).toEqual({ moves: 3, scrollTop: 3000, loaded: true, value: 'target' });
});

test('grid readiness follows a sheet change before the paint frame', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const first = await book.addSheet('Ready first', { rows: 100, columns: 2 });
    const second = await book.addSheet('Ready second', { rows: 100, columns: 2 });
    await second.range('A1').setValues([['second sheet']]);
    const grid = app.getGrid()!;
    grid.setSheet(first.id);
    await grid.ready();
    const original = window.requestAnimationFrame;
    let switched = false;
    window.requestAnimationFrame = (callback) =>
      original((time) => {
        if (!switched) {
          switched = true;
          grid.setSheet(second.id);
        }
        callback(time);
      });
    try {
      await grid.ready();
      return {
        switched,
        active: grid.activeSheetId,
        expected: second.id,
        loaded: book.isLoaded(second.id, 0, 0),
        value: book.peekValue(second.id, 0, 0),
      };
    } finally {
      window.requestAnimationFrame = original;
    }
  });
  expect(result).toMatchObject({ switched: true, loaded: true, value: 'second sheet' });
  expect(result.active).toBe(result.expected);
});

test('grid readiness follows a zoom change before the paint frame', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = await book.addSheet('Ready zoom', { rows: 150, columns: 2 });
    await sheet.range('A101').setValues([['zoom target']]);
    const grid = app.getGrid()!;
    grid.setSheet(sheet.id);
    await grid.ready();
    grid.scroller.scrollTop = 1500;
    await grid.ready();
    const original = window.requestAnimationFrame;
    let changed = false;
    window.requestAnimationFrame = (callback) =>
      original((time) => {
        if (!changed) {
          changed = true;
          grid.setZoom(0.5);
        }
        callback(time);
      });
    try {
      await grid.ready();
      return {
        changed,
        loaded: book.isLoaded(sheet.id, 100, 0),
        value: book.peekValue(sheet.id, 100, 0),
      };
    } finally {
      window.requestAnimationFrame = original;
    }
  });
  expect(result).toEqual({ changed: true, loaded: true, value: 'zoom target' });
});

test('grid readiness follows a committed edit before the paint frame', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const book = app.getWorkbook();
    const sheet = await book.addSheet('Ready revision', { rows: 100, columns: 2 });
    await sheet.range('A1').setValues([['old']]);
    const grid = app.getGrid()!;
    grid.setSheet(sheet.id);
    await grid.ready();
    const original = window.requestAnimationFrame;
    let changed = false;
    window.requestAnimationFrame = (callback) =>
      original((time) => {
        if (changed) return callback(time);
        changed = true;
        void sheet
          .range('A1')
          .setValues([['new']])
          .then(() => callback(time));
      });
    try {
      await grid.ready();
      return {
        changed,
        loaded: book.isLoaded(sheet.id, 0, 0),
        value: book.peekValue(sheet.id, 0, 0),
      };
    } finally {
      window.requestAnimationFrame = original;
    }
  });
  expect(result).toEqual({ changed: true, loaded: true, value: 'new' });
});

test('disposing a grid during the paint wait settles readiness', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const grid = window.opensheet.getGrid()!;
    await grid.ready();
    const original = window.requestAnimationFrame;
    let disposed = false;
    window.requestAnimationFrame = (callback) =>
      original((time) => {
        if (!disposed) {
          disposed = true;
          grid.dispose();
        }
        callback(time);
      });
    try {
      await grid.ready();
      return disposed;
    } finally {
      window.requestAnimationFrame = original;
    }
  });
  expect(result).toBe(true);
});

for (const existing of [false, true])
  test(`disposing a grid settles readiness while ${existing ? 'existing' : 'new'} prefetch is stalled`, async ({
    page,
  }) => {
    await page.goto('/#budget');
    await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
    const result = await page.evaluate(async (existingFetch) => {
      const app = window.opensheet;
      const book = app.getWorkbook();
      const grid = app.getGrid()!;
      await grid.ready();
      while ((grid as any).fetching) await (grid as any).fetching;
      const original = book.prefetch.bind(book);
      let start!: () => void;
      let rejectRead!: (error: Error) => void;
      const entered = new Promise<void>((resolve) => (start = resolve));
      const stalled = new Promise<void>((_resolve, reject) => (rejectRead = reject));
      const errors: string[] = [];
      (grid as any).options.onError = (error: Error) => errors.push(error.message);
      book.prefetch = () => {
        start();
        return stalled;
      };
      try {
        if (existingFetch) {
          void (grid as any).prepareViewport();
          await entered;
        }
        const ready = grid.ready();
        await entered;
        grid.dispose();
        const settled = await Promise.race([
          ready.then(() => true),
          new Promise<false>((resolve) => setTimeout(() => resolve(false), 500)),
        ]);
        rejectRead(new Error('late prefetch failure'));
        await new Promise((resolve) => setTimeout(resolve, 0));
        return { settled, errors };
      } finally {
        book.prefetch = original;
      }
    }, existing);
    expect(result).toEqual({ settled: true, errors: [] });
  });

test('completed readiness calls release their disposal listeners', async ({ page }) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async () => {
    const grid = window.opensheet.getGrid()!;
    await grid.ready();
    const signal: AbortSignal = (grid as any).abort.signal;
    const add = signal.addEventListener.bind(signal);
    const remove = signal.removeEventListener.bind(signal);
    let added = 0,
      removed = 0;
    signal.addEventListener = ((
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: any,
    ) => {
      if (type === 'abort') added++;
      add(type, listener, options);
    }) as typeof signal.addEventListener;
    signal.removeEventListener = ((
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: any,
    ) => {
      if (type === 'abort') removed++;
      remove(type, listener, options);
    }) as typeof signal.removeEventListener;
    try {
      for (let i = 0; i < 20; i++) await Promise.all([grid.ready(), grid.ready()]);
      return { added, removed };
    } finally {
      signal.addEventListener = add;
      signal.removeEventListener = remove;
    }
  });
  expect(result.added).toBeGreaterThan(0);
  expect(result.removed).toBe(result.added);
});
