import { build } from 'esbuild';
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { cpus, release } from 'node:os';
import { resolve } from 'node:path';
import { sourceFingerprint } from './benchmark-fingerprint.mjs';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/clipboard-parser-benchmark.mjs NEW_REPORT');
const path = resolve('packages/renderer/src/clipboard.ts');
const current = readFileSync(path, 'utf8');
const changes = [
  ['    closedQuote = false,\n', ''],
  ['    closedQuote = false;\n', ''],
  [
    "    if (closedQuote && c !== '\\t' && c !== '\\n' && c !== '\\r')\n      throw new OpenSheetError('INVALID_ARGUMENT', 'Unexpected text after closing quote');\n",
    '',
  ],
  [
    "    if (c === '\"' && (quoted || (!cell && !fieldStarted))) {",
    "    if (c === '\"' && (quoted || !cell)) {",
  ],
  ['        if (!quoted) closedQuote = true;\n', ''],
];
let previous = current;
for (const [from, to] of changes) {
  if (previous.split(from).length !== 2) throw new Error('Cannot reconstruct previous parser');
  previous = previous.replace(from, to);
}
const scripts = {},
  hashes = {};
for (const [variant, contents] of Object.entries({ previous, current })) {
  const result = await build({
    entryPoints: [path],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: `Clipboard_${variant}`,
    write: false,
    plugins: [
      {
        name: 'variant',
        setup(builder) {
          builder.onLoad({ filter: /renderer\/src\/clipboard\.ts$/ }, () => ({
            contents,
            loader: 'ts',
            resolveDir: resolve('packages/renderer/src'),
          }));
        },
      },
    ],
  });
  scripts[variant] = result.outputFiles[0].text;
  hashes[variant] = createHash('sha256').update(scripts[variant]).digest('hex');
}
const report = {
  startedAt: new Date().toISOString(),
  source: sourceFingerprint(),
  hashes,
  node: process.version,
  cpu: cpus()[0]?.model,
  platform: `${process.platform}/${process.arch}/${release()}`,
  scope:
    'Headless browser main-thread parser microbenchmark. Six alternating pairs for each of two valid 400×200 TSV fixtures; every resulting cell checked outside the timed interval. Previous reconstructs the parser before the closed-quote guard. Excludes clipboard events, commands, storage, paint, invalid inputs, peak memory, other devices and full application latency.',
  results: [],
};
const save = (flag = 'w') =>
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag });
save('wx');
for (const [engine, launcher] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await launcher.launch();
  const result = { engine, version: browser.version(), rounds: [] };
  report.results.push(result);
  try {
    const page = await browser.newPage();
    for (const script of Object.values(scripts)) await page.addScriptTag({ content: script });
    for (let round = 0; round < 6; round++) {
      const samples = await page.evaluate((round) => {
        const output = [];
        for (const variant of round % 2 ? ['current', 'previous'] : ['previous', 'current']) {
          for (const quoted of round % 2 ? [true, false] : [false, true]) {
            const cell = quoted ? '"a\tb"' : 'alpha';
            const expected = quoted ? 'a\tb' : 'alpha';
            const line = Array(200).fill(cell).join('\t');
            const input = Array(400).fill(line).join('\r\n');
            const parser = globalThis[`Clipboard_${variant}`].parseClipboardText;
            const start = performance.now();
            const rows = parser(input);
            const ms = performance.now() - start;
            if (
              rows.length !== 400 ||
              rows.some((row) => row.length !== 200 || row.some((value) => value !== expected))
            )
              throw new Error('Incorrect parsed matrix');
            output.push({ variant, quoted, ms, cells: 80000, inputLength: input.length });
          }
        }
        return output;
      }, round);
      result.rounds.push({ round: round + 1, samples });
      save();
    }
  } finally {
    await browser.close();
    save();
  }
}
report.finishedAt = new Date().toISOString();
save();
