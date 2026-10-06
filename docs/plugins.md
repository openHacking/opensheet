# Writing OpenSheet plugins

The current SDK apiVersion is `0.1.0` / `^0.1.0`. Install plugins with `app.use` after creating or loading a workbook.

```ts
import { definePlugin } from '@opensheetjs/plugin-sdk';

const sumPlugin = definePlugin({
  id: 'example.sum',
  version: '0.1.0',
  apiVersion: '^0.1.0',
  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],
  setup(ctx) {
    ctx.ui.toolbar.add({
      id: 'example.sum.button',
      label: 'Σ Sum',
      run() {
        const selection = ctx.selection.get();
        if (!selection) return;
        const values = ctx.workbook.readRange(selection).values.flat();
        const total = values.reduce<number>(
          (sum, value) => sum + (typeof value === 'number' ? value : 0),
          0,
        );
        ctx.ui.notify(String(total));
      },
    });
    return { api: { name: 'Selection sum' } };
  },
});
app.use(sumPlugin);
```

| Capability     | Available operations                              |
| -------------- | ------------------------------------------------- |
| selection.read | ctx.selection.get                                 |
| workbook.read  | ctx.workbook.getSnapshot / readRange              |
| workbook.write | ctx.workbook.execute using standard core commands |
| ui.toolbar     | ctx.ui.toolbar.add / notify                       |
| events         | ctx.onCommit                                      |

`requires: ['other.plugin']` declares dependencies. Batch installation sorts dependencies and rejects missing dependencies, cycles, duplicate IDs and incompatible versions. Setup failures clean registered resources; batch failures uninstall newly installed plugins. Business mutations during setup are not rolled back, so setup should register functionality rather than change data.

Toolbar and commit registrations belong to the plugin and are cleaned in reverse order. Use `ctx.addCleanup` or returned `dispose` for custom resources. Commit callbacks cannot synchronously mutate the workbook; start a later transaction from a user action or guarded microtask. `app.plugins.get<T>(id)` retrieves plugin APIs. Removal is rejected while dependents exist. Loading a workbook reinstalls active plugins against the new instance.

Capabilities are not a sandbox. Plugins share host JavaScript permissions and must be trusted. Custom editors, formula functions, shortcut registries, storage providers and untrusted-plugin isolation are not supported extension points in 0.1.

Start with headless tests for pure operations. Route writes through `ctx.workbook.execute` and check absent selections and error values. See the [command schema](../schemas/commands.schema.json) and [compile-checked example](../examples/plugin.ts).
