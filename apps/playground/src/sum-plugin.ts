import { definePlugin } from 'opensheet';
export const sumPlugin = definePlugin({
  id: 'example.selection-sum',
  version: '0.1.0',
  apiVersion: '^0.1.0',
  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],
  setup(ctx) {
    ctx.ui.toolbar.add({
      id: 'example.selection-sum.run',
      label: 'Σ Sum',
      async run() {
        const s = ctx.selection.get();
        if (s) {
          const values = (await ctx.workbook.readRange(s)).values.flat();
          const total = values.reduce<number>((sum, v) => sum + (typeof v === 'number' ? v : 0), 0);
          ctx.ui.notify(`Selected numbers sum to ${total.toLocaleString()}`);
        }
      },
    });
  },
});
