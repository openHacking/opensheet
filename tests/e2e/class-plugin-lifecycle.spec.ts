import { expect, test } from '@playwright/test';

test('class-based plugins retain private receiver state across workbook replacement', async ({
  page,
}) => {
  await page.goto('/');
  await expect(page.locator('#scene-select')).toBeEnabled();
  const result = await page.evaluate(async () => {
    const app = window.opensheet;
    let disposed = 0;
    let clicks = 0;
    // Keep native private fields inside the page: Playwright serializes this callback
    // without the transpiler helpers used for private fields in the test process.
    const ClassPlugin = eval(`(class ClassPlugin {
      #installs = 0;
      get id() {
        return 'example.class-plugin';
      }
      get version() {
        return '0.1.0';
      }
      get apiVersion() {
        return '^0.1.0';
      }
      capabilities = ['ui.toolbar'];
      setup(ctx) {
        const count = ++this.#installs;
        ctx.ui.toolbar.add({
          id: 'class-plugin',
          label: 'Class plugin',
          run() {
            clicks += count;
          },
        });
        return {
          api: { count },
          dispose() {
            disposed++;
          },
        };
      }
    })`);
    const plugin = new ClassPlugin();
    app.use(plugin);
    const first = app.plugins.get('example.class-plugin');
    (document.querySelector('[data-action="class-plugin"]') as HTMLButtonElement).click();
    await app.open(app.getWorkbook().id);
    const second = app.plugins.get('example.class-plugin');
    const actions = document.querySelectorAll('[data-action="class-plugin"]').length;
    (document.querySelector('[data-action="class-plugin"]') as HTMLButtonElement).click();
    // Toolbar actions are dispatched in a microtask; allow it to run before uninstall.
    await Promise.resolve();
    app.plugins.remove('example.class-plugin');
    return {
      first,
      second,
      actions,
      clicks,
      disposed,
      removed: !document.querySelector('[data-action="class-plugin"]'),
    };
  });
  expect(result).toEqual({
    first: { count: 1 },
    second: { count: 2 },
    actions: 1,
    clicks: 3,
    disposed: 2,
    removed: true,
  });
});
