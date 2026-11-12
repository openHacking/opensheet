// Run against a bundled SDK module: node scripts/plugin-registry-benchmark.mjs /absolute/path.mjs
import { pathToFileURL } from 'node:url';
import { performance } from 'node:perf_hooks';

const modulePath = process.argv[2];
if (!modulePath?.startsWith('/')) throw new Error('Pass an absolute SDK module path');
const { PluginRegistry } = await import(pathToFileURL(modulePath).href);
const host = {
  onError(error) {
    throw error;
  },
};
const results = [];
for (const count of [2000, 5000, 10000]) {
  const samples = [];
  for (let run = 0; run < 3; run++) {
    let disposed = 0;
    const plugins = Array.from({ length: count }, (_, i) => ({
      id: `example.plugin-${i}`,
      version: '0.1.0',
      apiVersion: '^0.1.0',
      requires: i ? [`example.plugin-${i - 1}`] : [],
      capabilities: [],
      setup() {
        return {
          dispose() {
            disposed++;
          },
        };
      },
    }));
    const registry = new PluginRegistry(host);
    const start = performance.now();
    registry.use(plugins);
    const installed = performance.now();
    registry.dispose();
    const done = performance.now();
    if (disposed !== count || registry.definitions().length) throw new Error('Incomplete disposal');
    samples.push({ installMs: installed - start, disposeMs: done - installed });
  }
  results.push({ count, samples });
}
console.log(
  JSON.stringify(
    { node: process.version, platform: process.platform, arch: process.arch, results },
    null,
    2,
  ),
);
