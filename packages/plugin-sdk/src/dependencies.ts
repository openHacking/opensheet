import { OpenSheetError } from '@opensheetjs/core';
import type { Plugin } from './types.js';
export function orderPlugins(plugins: Plugin[], has: (id: string) => boolean): Plugin[] {
  const pending = new Map<string, Plugin>();
  for (const p of plugins) {
    if (pending.has(p.id) || has(p.id))
      throw new OpenSheetError('PLUGIN_CONFLICT', `Duplicate plugin ${p.id}`);
    if (!/^[a-z][\w.-]+$/.test(p.id) || !['0.1.0', '^0.1.0'].includes(p.apiVersion))
      throw new OpenSheetError('PLUGIN_CONFLICT', `Incompatible plugin ${p.id}`);
    pending.set(p.id, p);
  }
  const order: Plugin[] = [],
    visiting = new Set<string>(),
    visited = new Set<string>();
  for (const root of plugins) {
    if (visited.has(root.id)) continue;
    const stack: Array<{ plugin: Plugin; next: number }> = [{ plugin: root, next: 0 }];
    visiting.add(root.id);
    while (stack.length) {
      const frame = stack[stack.length - 1];
      const dependencies = frame.plugin.requires ?? [];
      if (frame.next === dependencies.length) {
        stack.pop();
        visiting.delete(frame.plugin.id);
        visited.add(frame.plugin.id);
        order.push(frame.plugin);
        continue;
      }
      const id = dependencies[frame.next++];
      if (has(id)) continue;
      const dep = pending.get(id);
      if (!dep) throw new OpenSheetError('PLUGIN_CONFLICT', `Missing dependency ${id}`);
      if (visited.has(id)) continue;
      if (visiting.has(id)) throw new OpenSheetError('PLUGIN_CONFLICT', 'Plugin dependency cycle');
      visiting.add(id);
      stack.push({ plugin: dep, next: 0 });
    }
  }
  return order;
}
