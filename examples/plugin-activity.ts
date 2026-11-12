import { definePlugin } from '@opensheetjs/plugin-sdk';

/** Counts commits while this installation is active; reopening starts a new count. */
export const activityCounter = definePlugin({
  id: 'example.activity-counter',
  version: '0.1.0',
  apiVersion: '^0.1.0',
  capabilities: ['events', 'ui.toolbar'],
  setup(ctx) {
    let commits = 0;
    ctx.onCommit(() => {
      commits++;
    });
    ctx.ui.toolbar.add({
      id: 'example.activity-counter.show',
      label: 'Changes',
      run() {
        ctx.ui.notify(`${commits} committed changes in this session`);
      },
    });
    return { api: { getCount: () => commits } };
  },
});
