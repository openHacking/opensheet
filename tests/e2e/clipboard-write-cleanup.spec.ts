import { expect, test } from '@playwright/test';

test('clipboard denial does not wait for a stalled structured read and observes its later failure', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (url) => {
    const { writeClipboard } = await import(/* @vite-ignore */ url);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const before = { revision: book.revision, value: await sheet.range('A1').getValues() };
    const originalRead = book.readCells;
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const originalItem = Object.getOwnPropertyDescriptor(window, 'ClipboardItem');
    const errors: string[] = [];
    const onUnhandled = (event: PromiseRejectionEvent) => {
      errors.push(String(event.reason));
      event.preventDefault();
    };
    window.addEventListener('unhandledrejection', onUnhandled);
    let rejectRead!: (reason: unknown) => void;
    const pending = new Promise<never>((_resolve, reject) => {
      rejectRead = reject;
    });
    book.readCells = async () => pending;
    const denied = new DOMException('Clipboard denied', 'NotAllowedError');
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async () => {
          throw denied;
        },
      },
    });
    Object.defineProperty(window, 'ClipboardItem', {
      configurable: true,
      value: class {
        constructor(_data: unknown) {}
      },
    });
    try {
      const operation = writeClipboard(sheet.range('A1')).then(
        () => 'success',
        (error: unknown) => (error === denied ? 'denied' : String(error)),
      );
      const first = await Promise.race([
        operation,
        new Promise<string>((resolve) => setTimeout(() => resolve('timeout'), 200)),
      ]);
      rejectRead(new Error('Late read failure'));
      const final = await operation;
      await new Promise((resolve) => setTimeout(resolve, 50));
      return {
        first,
        final,
        errors,
        before,
        after: { revision: book.revision, value: await sheet.range('A1').getValues() },
      };
    } finally {
      book.readCells = originalRead;
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else delete (navigator as any).clipboard;
      if (originalItem) Object.defineProperty(window, 'ClipboardItem', originalItem);
      else delete (window as any).ClipboardItem;
      window.removeEventListener('unhandledrejection', onUnhandled);
    }
  }, `/@fs${process.cwd()}/packages/renderer/src/copy-clipboard.ts`);
  expect(result.first).toBe('denied');
  expect(result.final).toBe('denied');
  expect(result.errors).toEqual([]);
  expect(result.after).toEqual(result.before);
});

test('an early clipboard success still waits for and reports a failed structured read', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (url) => {
    const { writeClipboard } = await import(/* @vite-ignore */ url);
    const book = window.opensheet.getWorkbook();
    const sheet = book.getSheets()[0];
    const originalRead = book.readCells;
    const originalClipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const originalItem = Object.getOwnPropertyDescriptor(window, 'ClipboardItem');
    let rejectRead!: (reason: unknown) => void;
    const pending = new Promise<never>((_resolve, reject) => {
      rejectRead = reject;
    });
    const failure = new Error('Structured read failed');
    book.readCells = async () => pending;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { write: async () => {} },
    });
    Object.defineProperty(window, 'ClipboardItem', {
      configurable: true,
      value: class {
        constructor(_data: unknown) {}
      },
    });
    try {
      const operation = writeClipboard(sheet.range('A1')).then(
        () => 'success',
        (error: unknown) => (error === failure ? 'read failed' : String(error)),
      );
      const before = await Promise.race([
        operation,
        new Promise<string>((resolve) => setTimeout(() => resolve('pending'), 100)),
      ]);
      rejectRead(failure);
      return { before, after: await operation };
    } finally {
      book.readCells = originalRead;
      if (originalClipboard) Object.defineProperty(navigator, 'clipboard', originalClipboard);
      else delete (navigator as any).clipboard;
      if (originalItem) Object.defineProperty(window, 'ClipboardItem', originalItem);
      else delete (window as any).ClipboardItem;
    }
  }, `/@fs${process.cwd()}/packages/renderer/src/copy-clipboard.ts`);
  expect(result).toEqual({ before: 'pending', after: 'read failed' });
});

test('grid shows clipboard denial while the underlying copy read remains stalled', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const before = await page.evaluate(() => {
    const book = window.opensheet.getWorkbook();
    const originalRead = book.readCells;
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const item = Object.getOwnPropertyDescriptor(window, 'ClipboardItem');
    let rejectRead!: (reason: unknown) => void;
    const pending = new Promise<never>((_resolve, reject) => {
      rejectRead = reject;
    });
    book.readCells = async () => pending;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        write: async () => {
          throw new DOMException('Clipboard denied', 'NotAllowedError');
        },
      },
    });
    Object.defineProperty(window, 'ClipboardItem', {
      configurable: true,
      value: class {
        constructor(_data: unknown) {}
      },
    });
    (window as any).releaseCopyRead = () => {
      rejectRead(new Error('Late read failure'));
      book.readCells = originalRead;
      if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
      else delete (navigator as any).clipboard;
      if (item) Object.defineProperty(window, 'ClipboardItem', item);
      else delete (window as any).ClipboardItem;
    };
    return book.revision;
  });
  await page.getByRole('grid', { exact: true }).dispatchEvent('copy');
  await expect(page.locator('.os-toast')).toContainText('Clipboard denied');
  expect(await page.evaluate(() => window.opensheet.getWorkbook().revision)).toBe(before);
  await page.evaluate(() => (window as any).releaseCopyRead());
});

test('clipboard item constructor failure returns before a stalled read settles', async ({
  page,
}) => {
  await page.goto('/#budget');
  await expect(page.locator('#scene-title')).toHaveText('Budget Simulator');
  const result = await page.evaluate(async (url) => {
    const { writeClipboard } = await import(/* @vite-ignore */ url);
    const book = window.opensheet.getWorkbook();
    const originalRead = book.readCells;
    const clipboard = Object.getOwnPropertyDescriptor(navigator, 'clipboard');
    const item = Object.getOwnPropertyDescriptor(window, 'ClipboardItem');
    const failure = new Error('Item construction failed');
    let rejectRead!: (reason: unknown) => void;
    const pending = new Promise<never>((_resolve, reject) => {
      rejectRead = reject;
    });
    book.readCells = async () => pending;
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: { write: async () => {} },
    });
    Object.defineProperty(window, 'ClipboardItem', {
      configurable: true,
      value: class {
        constructor(_data: unknown) {
          throw failure;
        }
      },
    });
    try {
      const operation = writeClipboard(book.getSheets()[0].range('A1')).then(
        () => 'success',
        (error: unknown) => (error === failure ? 'failed' : String(error)),
      );
      const first = await Promise.race([
        operation,
        new Promise<string>((resolve) => setTimeout(() => resolve('timeout'), 200)),
      ]);
      rejectRead(new Error('Late failure'));
      return { first, final: await operation };
    } finally {
      book.readCells = originalRead;
      if (clipboard) Object.defineProperty(navigator, 'clipboard', clipboard);
      else delete (navigator as any).clipboard;
      if (item) Object.defineProperty(window, 'ClipboardItem', item);
      else delete (window as any).ClipboardItem;
    }
  }, `/@fs${process.cwd()}/packages/renderer/src/copy-clipboard.ts`);
  expect(result).toEqual({ first: 'failed', final: 'failed' });
});
