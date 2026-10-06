import { definePlugin } from '@opensheetjs/plugin-sdk';
export const selectionSum = definePlugin({
  id: 'example.sum',
  version: '0.1.0',
  apiVersion: '^0.1.0',
  capabilities: ['selection.read', 'workbook.read', 'ui.toolbar'],
  setup(ctx) {
    ctx.ui.toolbar.add({
      id: 'example.sum.button',
      label: 'Sum',
      run() {
        const selected = ctx.selection.get();
        if (!selected) return;
        const sum = ctx.workbook
          .readRange(selected)
          .values.flat()
          .reduce<number>((n, v) => n + (typeof v === 'number' ? v : 0), 0);
        ctx.ui.notify(String(sum));
      },
    });
  },
});
