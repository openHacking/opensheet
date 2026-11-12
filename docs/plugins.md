> Workbook reads/commands are asynchronous. Await `readRange`, `toJSON` and `execute`; toolbar actions may return promises. `toJSON` returns a complete WorkbookFile; `getMetadata` is lightweight, so use bounded range reads during ordinary interaction.

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
      async run() {
        const selection = ctx.selection.get();
        if (!selection) return;
        const values = (await ctx.workbook.readRange(selection)).values.flat();
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
| workbook.read  | ctx.workbook.getMetadata / toJSON / readRange     |
| workbook.write | ctx.workbook.execute using standard core commands |
| ui.toolbar     | ctx.ui.toolbar.add / notify                       |
| events         | ctx.onCommit                                      |

`requires: ['other.plugin']` declares dependencies. Batch installation sorts dependencies and rejects missing dependencies, cycles, duplicate IDs and incompatible versions. Setup failures clean registered resources; batch failures uninstall newly installed plugins. Business mutations during setup are not rolled back, so setup should register functionality rather than change data.

Toolbar and commit registrations belong to the plugin and are cleaned in reverse order. Use `ctx.addCleanup` or returned `dispose` for custom resources. Commit callbacks cannot synchronously mutate the workbook; start a later transaction from a user action or guarded microtask. `app.plugins.get<T>(id)` retrieves plugin APIs. Removal is rejected while dependents exist. Loading a workbook reinstalls active plugins against the new instance.

Capabilities are not a sandbox. Plugins share host JavaScript permissions and must be trusted. Custom editors, formula functions, shortcut registries, storage providers and untrusted-plugin isolation are not supported extension points in 0.1.

Start with headless tests for pure operations. Route writes through `ctx.workbook.execute` and check absent selections and error values. The [selection-sum action](../examples/plugin.ts) reads a bounded selection; the [activity counter](../examples/plugin-activity.ts) registers a commit listener and a toolbar action. Both use the public SDK entry point, declare API compatibility with `^0.1.0`, and are covered by the repository's Apache-2.0 license. The counter is per installation: replacing a workbook reinstalls the plugin and resets the count. Neither example is an independent adoption claim.

The preview `checkPluginConformance` helper accepts a plugin factory and a disposable `PluginHost`, then reports named lifecycle and registry checks. It executes the plugin and cannot detect arbitrary plugin-owned leaks or await asynchronous cleanup; see the [SDK guide](../packages/plugin-sdk/README.md) for its scope. See also the [command schema](../schemas/commands.schema.json).
