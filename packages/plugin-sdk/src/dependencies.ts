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
  const visit = (p: Plugin) => {
    if (visited.has(p.id)) return;
    if (visiting.has(p.id)) throw new OpenSheetError('PLUGIN_CONFLICT', 'Plugin dependency cycle');
    visiting.add(p.id);
    for (const id of p.requires ?? []) {
      if (has(id)) continue;
      const dep = pending.get(id);
      if (!dep) throw new OpenSheetError('PLUGIN_CONFLICT', `Missing dependency ${id}`);
      visit(dep);
    }
    visiting.delete(p.id);
    visited.add(p.id);
    order.push(p);
  };
  plugins.forEach(visit);
  return order;
}
