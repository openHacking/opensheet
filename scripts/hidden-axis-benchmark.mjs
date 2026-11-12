import { build } from 'esbuild';
import { writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { chromium, firefox, webkit } from '@playwright/test';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/hidden-axis-benchmark.mjs NEW_REPORT');
const bundle = await build({
  entryPoints: ['packages/renderer/src/layout.ts'],
  bundle: true,
  platform: 'browser',
  format: 'iife',
  globalName: 'GridLayoutBundle',
  write: false,
});
const report = {
  startedAt: new Date().toISOString(),
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  source: sourceFingerprint(),
  scope:
    'Paired synthetic browser enumeration only; excludes layout rebuild, canvas paint, workbook, memory and reference-device acceptance. Old linear loop is reconstructed in the same page. Repeated batches alternate order and verify every returned index.',
  browsers: {},
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [name, engine] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await engine.launch({ headless: true });
  try {
    const page = await browser.newPage();
    await page.addScriptTag({ content: bundle.outputFiles[0].text });
    const result = await page.evaluate(() => {
      const layout = new globalThis.GridLayoutBundle.GridLayout();
      const linear = (offsets, scroll, extent, frozen) => {
        const out = [];
        for (let i = 0; i < Math.min(frozen, offsets.length - 1) && offsets[i] < extent; i++)
          if (offsets[i + 1] > offsets[i]) out.push(i);
        const start = Math.max(frozen, layout.index(offsets, scroll + offsets[frozen]));
        for (let i = start; i < offsets.length - 1 && offsets[i] - scroll < extent; i++)
          if (offsets[i + 1] > offsets[i]) out.push(i);
        return out;
      };
      const cases = [
        {
          name: 'visible',
          sizes: Array(100001).fill(30),
          frozen: 2,
          extent: 1000,
          iterations: 50000,
        },
        {
          name: 'alternating',
          sizes: Array.from({ length: 100001 }, (_, i) => (i % 2 ? 0 : 30)),
          frozen: 2,
          extent: 1000,
          iterations: 50000,
        },
        {
          name: 'frozen-hidden',
          sizes: [...Array(100000).fill(0), 30],
          frozen: 100000,
          extent: 1000,
          iterations: 200,
        },
      ];
      const results = [];
      for (const fixture of cases) {
        const offsets = [0];
        for (const size of fixture.sizes) offsets.push(offsets[offsets.length - 1] + size);
        const expected = linear(offsets, 0, fixture.extent, fixture.frozen);
        const variants = { linear, optimized: (...args) => layout.visiblePositions(...args) };
        for (let round = 0; round < 6; round++) {
          for (const name of round % 2 ? ['optimized', 'linear'] : ['linear', 'optimized']) {
            const fn = variants[name];
            let verified = 0;
            const start = performance.now();
            for (let i = 0; i < fixture.iterations; i++) {
              const actual = fn(offsets, 0, fixture.extent, fixture.frozen);
              if (actual.length !== expected.length)
                throw new Error(`${fixture.name} ${name} length`);
              for (let j = 0; j < actual.length; j++)
                if (actual[j] !== expected[j])
                  throw new Error(`${fixture.name} ${name} index ${j}`);
              verified += actual.length;
            }
            results.push({
              fixture: fixture.name,
              round,
              variant: name,
              ms: performance.now() - start,
              iterations: fixture.iterations,
              verified,
            });
          }
        }
      }
      return results;
    });
    report.browsers[name] = { version: browser.version(), samples: result };
    save();
    console.log(name, 'completed', result.length, 'samples');
  } finally {
    await browser.close();
  }
}
report.finishedAt = new Date().toISOString();
save();
