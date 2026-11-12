import { expect, test } from '@playwright/test';

test('removal keeps the original returned disposer and receiver', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(() => {
    const calls: string[] = [];
    const resource = {
      marker: 'original',
      dispose() {
        calls.push(this.marker);
      },
    };
    window.opensheet.use({
      id: 'example.disposer',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [],
      setup() {
        return resource;
      },
    });
    resource.dispose = () => calls.push('replacement');
    window.opensheet.plugins.remove('example.disposer');
    return calls;
  });
  expect(result).toEqual(['original']);
});

test('rejects asynchronous setup and removes resources before its continuation', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const errors: string[] = [];
    const unhandled: string[] = [];
    const listener = (event: PromiseRejectionEvent) => {
      unhandled.push(String(event.reason));
      event.preventDefault();
    };
    window.addEventListener('unhandledrejection', listener);
    app.on('plugin:error', (error) => errors.push((error as Error).message));
    let message = '';
    try {
      app.use({
        id: 'example.async-setup',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: ['ui.toolbar'],
        setup: (async (ctx: any) => {
          ctx.ui.toolbar.add({ id: 'async-first', label: 'First', run() {} });
          await Promise.resolve();
          ctx.ui.toolbar.add({ id: 'async-late', label: 'Late', run() {} });
        }) as any,
      });
    } catch (error) {
      message = (error as Error).message;
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
    window.removeEventListener('unhandledrejection', listener);
    return {
      message,
      errors,
      unhandled,
      installed: app.plugins.has('example.async-setup'),
      first: !!document.querySelector('[data-action="async-first"]'),
      late: !!document.querySelector('[data-action="async-late"]'),
    };
  });
  expect(result.message).toContain('setup must be synchronous');
  expect(result.errors).toEqual(
    expect.arrayContaining([expect.stringContaining('no longer active')]),
  );
  expect(result.unhandled).toEqual([]);
  expect(result.installed).toBe(false);
  expect(result.first).toBe(false);
  expect(result.late).toBe(false);
});

test('plugins installed through the registry survive workbook replacement in dependency order', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const order: string[] = [];
    const dependency = {
      id: 'example.direct-dependency',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [] as never[],
      setup() {
        order.push('dependency');
      },
    };
    const dependent = {
      id: 'example.direct-dependent',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: ['example.direct-dependency'],
      capabilities: [] as never[],
      setup() {
        order.push('dependent');
      },
    };
    app.plugins.use([dependent, dependency]);
    const copy = app.plugins.definitions();
    const copied = copy.find((item) => item.id === 'example.direct-dependent')!;
    copied.requires!.length = 0;
    dependency.id = 'example.changed';
    await app.load(await app.getWorkbook().toJSON());
    let dependencyError = '';
    try {
      app.plugins.remove('example.direct-dependency');
    } catch (error) {
      dependencyError = (error as Error).message;
    }
    return {
      order,
      dependency: app.plugins.has('example.direct-dependency'),
      dependent: app.plugins.has('example.direct-dependent'),
      changed: app.plugins.has('example.changed'),
      dependencyError,
    };
  });
  expect(result.order).toEqual(['dependency', 'dependent', 'dependency', 'dependent']);
  expect(result.dependency).toBe(true);
  expect(result.dependent).toBe(true);
  expect(result.changed).toBe(false);
  expect(result.dependencyError).toContain('depend');
});

for (const format of ['json', 'binary'] as const) {
  test(`reports plugin failure after committed ${format} import without rejecting the import`, async ({
    page,
  }) => {
    await page.goto('/');
    await expect(page.locator('#scene-select')).toBeEnabled();
    const result = await page.evaluate(async (format) => {
      const app = window.opensheet;
      const errors: string[] = [];
      let installs = 0;
      app.on('plugin:error', (error) => errors.push((error as Error).message));
      app.use({
        id: 'example.import-restore',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: [],
        setup() {
          if (++installs > 1) throw new Error('import plugin restore failed');
        },
      });
      const book = app.getWorkbook();
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[123]]);
      let saved: string | Blob;
      if (format === 'json') saved = JSON.stringify(await book.toJSON());
      else {
        const parts: Uint8Array[] = [];
        for await (const chunk of book.streamBinary()) parts.push(chunk);
        saved = new Blob(parts as BlobPart[]);
      }
      await book
        .getSheets()[0]
        .range('A1')
        .setValues([[456]]);
      if (format === 'json') await app.loadJSON(saved as string);
      else await app.loadBinary(saved as Blob);
      const current = app.getWorkbook();
      const loaded = await current.getSheets()[0].range('A1').getValues();
      await app.open(current.id);
      return {
        loaded,
        reopened: await app.getWorkbook().getSheets()[0].range('A1').getValues(),
        installs,
        installed: app.plugins.has('example.import-restore'),
        errors,
      };
    }, format);
    expect(result.loaded).toEqual([[123]]);
    expect(result.reopened).toEqual([[123]]);
    expect(result.installs).toBe(2);
    expect(result.installed).toBe(false);
    expect(result.errors).toContain('import plugin restore failed');
  });
}

test('reports plugin restoration failures without misreporting workbook replacement', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let badInstalls = 0,
      goodInstalls = 0,
      dependentInstalls = 0;
    const errors: string[] = [];
    app.on('plugin:error', (error) => errors.push((error as Error).message));
    const bad = {
      id: 'example.bad-restore',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: ['ui.toolbar'] as 'ui.toolbar'[],
      setup(ctx: any) {
        badInstalls++;
        ctx.ui.toolbar.add({ id: 'bad-restore', label: 'Bad restore', run() {} });
        if (badInstalls > 1) throw new Error('plugin restore failed');
      },
    };
    const dependent = {
      id: 'example.dependent-restore',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: ['example.bad-restore'],
      capabilities: [] as never[],
      setup() {
        dependentInstalls++;
      },
    };
    const good = {
      id: 'example.good-restore',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [] as never[],
      setup() {
        goodInstalls++;
      },
    };
    app.use([dependent, good, bad]);
    const old = app.getWorkbook();
    await old
      .getSheets()[0]
      .range('A1')
      .setValues([[123]]);
    const saved = await old.toJSON();
    await old
      .getSheets()[0]
      .range('A1')
      .setValues([[456]]);
    await app.load(saved);
    const current = app.getWorkbook();
    const loaded = await current.getSheets()[0].range('A1').getValues();
    const id = current.id;
    await app.open(id);
    return {
      changed: old.id !== id,
      loaded,
      reopened: await app.getWorkbook().getSheets()[0].range('A1').getValues(),
      badInstalls,
      goodInstalls,
      dependentInstalls,
      bad: app.plugins.has('example.bad-restore'),
      dependent: app.plugins.has('example.dependent-restore'),
      good: app.plugins.has('example.good-restore'),
      actionPresent: !!document.querySelector('[data-action="bad-restore"]'),
      errors,
    };
  });
  expect(result.changed).toBe(true);
  expect(result.loaded).toEqual([[123]]);
  expect(result.reopened).toEqual([[123]]);
  expect(result.badInstalls).toBe(2);
  expect(result.goodInstalls).toBe(3);
  expect(result.dependentInstalls).toBe(1);
  expect(result.bad).toBe(false);
  expect(result.dependent).toBe(false);
  expect(result.good).toBe(true);
  expect(result.actionPresent).toBe(false);
  expect(result.errors).toEqual(
    expect.arrayContaining([
      expect.stringContaining('plugin restore failed'),
      expect.stringContaining('could not restore dependency'),
    ]),
  );
});

test('workbook replacement reinstalls only the latest definition after uninstall and reinstall', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let oldInstalls = 0,
      newInstalls = 0;
    const definition = (setup: () => void) => ({
      id: 'example.reinstall',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: [] as never[],
      setup,
    });
    app.use(
      definition(() => {
        oldInstalls++;
      }),
    );
    app.plugins.remove('example.reinstall');
    app.use(
      definition(() => {
        newInstalls++;
      }),
    );
    await app.load(await app.getWorkbook().toJSON());
    return { oldInstalls, newInstalls, installed: app.plugins.has('example.reinstall') };
  });
  expect(result).toEqual({ oldInstalls: 1, newInstalls: 2, installed: true });
});

test('workbook replacement reinstalls the original plugin declaration', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    const contexts: any[] = [];
    const plugin = {
      id: 'example.stable',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: ['workbook.read'] as ('workbook.read' | 'selection.read')[],
      setup(ctx: any) {
        contexts.push(ctx);
      },
    };
    app.use(plugin);
    plugin.id = 'example.changed';
    plugin.capabilities.push('selection.read');
    await app.load(await app.getWorkbook().toJSON());
    let oldError = '',
      capabilityError = '';
    try {
      contexts[0].workbook.getMetadata();
    } catch (error) {
      oldError = (error as Error).message;
    }
    try {
      contexts[1].selection.get();
    } catch (error) {
      capabilityError = (error as Error).message;
    }
    return {
      count: contexts.length,
      original: app.plugins.has('example.stable'),
      changed: app.plugins.has('example.changed'),
      oldError,
      capabilityError,
    };
  });
  expect(result.count).toBe(2);
  expect(result.original).toBe(true);
  expect(result.changed).toBe(false);
  expect(result.oldError).toContain('no longer active');
  expect(result.capabilityError).toContain('did not declare');
});

test('rolls back a plugin that attempts reentrant installation', async ({ page }) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(() => {
    const app = window.opensheet;
    let error = '';
    try {
      app.use({
        id: 'example.outer',
        version: '0.1.0',
        apiVersion: '^0.1.0',
        capabilities: ['ui.toolbar'],
        setup(ctx) {
          ctx.ui.toolbar.add({ id: 'reentrant-check', label: 'Reentrant check', run() {} });
          app.use({
            id: 'example.inner',
            version: '0.1.0',
            apiVersion: '^0.1.0',
            capabilities: [],
            setup() {},
          });
        },
      });
    } catch (e) {
      error = (e as Error).message;
    }
    return {
      error,
      outer: app.plugins.has('example.outer'),
      inner: app.plugins.has('example.inner'),
      removed: !document.querySelector('[data-action="reentrant-check"]'),
    };
  });
  expect(result.error).toContain('mutation is already in progress');
  expect(result.outer).toBe(false);
  expect(result.inner).toBe(false);
  expect(result.removed).toBe(true);
});

test('uninstalling a plugin revokes its retained context and removes its toolbar action', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let context: any;
    let actions = 0;
    app.use({
      id: 'example.lifecycle-check',
      version: '0.1.0',
      apiVersion: '^0.1.0',
      capabilities: ['ui.toolbar', 'workbook.read'],
      setup(ctx) {
        context = ctx;
        ctx.ui.toolbar.add({
          id: 'lifecycle-check',
          label: 'Lifecycle check',
          run: () => {
            actions++;
          },
        });
      },
    });
    const button = document.querySelector<HTMLButtonElement>('[data-action="lifecycle-check"]')!;
    button.click();
    await new Promise((resolve) => setTimeout(resolve, 0));
    app.plugins.remove('example.lifecycle-check');
    button.click();
    let error = '';
    try {
      context.workbook.toJSON();
    } catch (e) {
      error = (e as Error).message;
    }
    return { actions, error, removed: !document.querySelector('[data-action="lifecycle-check"]') };
  });
  expect(result.actions).toBe(1);
  expect(result.error).toContain('no longer active');
  expect(result.removed).toBe(true);
});
