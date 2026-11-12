import { chromium, firefox, webkit } from '@playwright/test';
import { readFileSync, readdirSync, writeFileSync, mkdirSync, renameSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { dirname, resolve } from 'node:path';
import { arch, platform, cpus, totalmem, release } from 'node:os';
import { execFileSync } from 'node:child_process';
import { sourceFingerprint, localBuildFingerprint } from './benchmark-fingerprint.mjs';
import { startProcessMemorySampler } from './process-memory-sampler.mjs';
const option = (name, fallback) =>
  process.argv.find((arg) => arg.startsWith(`--${name}=`))?.slice(name.length + 3) ?? fallback;
const url = new URL(option('url', 'http://127.0.0.1:5173/'));
url.hash = 'performance';
const output = resolve(option('output', 'benchmarks/capacity-baseline.json'));
const timeout = Number(option('timeout-ms', '3600000'));
const budget = Number(option('budget-mib', '256'));
if (
  !Number.isSafeInteger(budget) ||
  budget < 1 ||
  budget > 8192 ||
  !Number.isFinite(timeout) ||
  timeout < 1000
)
  throw new Error('Invalid budget or timeout');
const sourceCommit = execFileSync('git', ['rev-parse', 'HEAD'], { encoding: 'utf8' }).trim();
const sourceDirty = Boolean(
  execFileSync('git', ['status', '--porcelain'], { encoding: 'utf8' }).trim(),
);
const cpuProfilePath = option('cpu-profile', '');
const browserName = option('browser', 'chromium');
const launchers = { chromium, firefox, webkit };
if (!Object.hasOwn(launchers, browserName)) throw new Error('Invalid browser');
if (cpuProfilePath && browserName !== 'chromium')
  throw new Error('CPU profiling requires Chromium');
if (process.argv.includes('--process-memory') && browserName !== 'chromium')
  throw new Error('Process memory sampling currently supports Chromium only');
const sourceDigest = sourceFingerprint();
const localBuildDigest = localBuildFingerprint();
let servedArtifacts;
async function verifyServedBuild() {
  const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
  const artifacts = [];
  for (const file of [
    'index.html',
    ...readdirSync('apps/playground/dist/assets').map((name) => `assets/${name}`),
  ]) {
    const response = await fetch(new URL(file, url));
    if (!response.ok) throw new Error(`Cannot fetch built artifact: ${file}`);
    const sha256 = hash(new Uint8Array(await response.arrayBuffer()));
    if (sha256 !== hash(readFileSync(`apps/playground/dist/${file}`)))
      throw new Error(`Served artifact differs from local build: ${file}`);
    artifacts.push({ file, sha256 });
  }
  return artifacts;
}
if (process.argv.includes('--verify-build')) servedArtifacts = await verifyServedBuild();
const browser = await launchers[browserName].launch();
const memorySampler = process.argv.includes('--process-memory')
  ? startProcessMemorySampler()
  : undefined;
let interval,
  profiler,
  saved = Promise.resolve(),
  previous = '';
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
  if (cpuProfilePath) {
    profiler = await page.context().newCDPSession(page);
    await profiler.send('Profiler.enable');
    await profiler.send('Profiler.start');
  }
  page.on('pageerror', (error) => process.stderr.write(`Page error: ${error.message}\n`));
  await page.goto(url.href);
  await page.locator('#scene-title').filter({ hasText: 'Performance Lab' }).waitFor();
  await page.waitForFunction(() => !document.querySelector('#scene-select').disabled);
  await page.locator('#perf-budget').fill(String(budget));
  const save = async () => {
    const download = page.waitForEvent('download');
    await page.locator('#perf-download').click();
    const file = await (await download).path(),
      result = JSON.parse(readFileSync(file, 'utf8'));
    Object.assign(result, {
      platform: platform(),
      arch: arch(),
      node: process.version,
      headless: true,
      engineFormat: 'binary-pages-schema3',
      browserVersion: browser.version(),
      browserEngine: browserName,
      sourceCommit,
      sourceDirty,
      sourceFingerprint: sourceDigest,
      localBuildFingerprint: localBuildDigest,
      ...(servedArtifacts ? { servedArtifacts } : {}),
      hardware: { cpu: cpus()[0]?.model, logicalCores: cpus().length, memoryBytes: totalmem() },
      osRelease: release(),
      powerMode: option('power-mode', 'unrecorded'),
      profiling:
        [
          cpuProfilePath ? 'main-document CPU sampling' : '',
          url.searchParams.has('trace-ops') ? 'worker operation timing' : '',
        ]
          .filter(Boolean)
          .join(' + ') || 'none',
      ...(memorySampler ? { processMemory: memorySampler.snapshot() } : {}),
    });
    mkdirSync(dirname(output), { recursive: true });
    writeFileSync(output + '.tmp', JSON.stringify(result, null, 2) + '\n');
    renameSync(output + '.tmp', output);
    return result;
  };
  const cells = option('cells', '');
  if (cells) {
    const count = Number(cells);
    if (!Number.isSafeInteger(count) || count < 10000 || count > 10000000 || count % 10000)
      throw new Error('Invalid cell count');
    await page.locator('#perf-cells').fill(cells);
  }
  await page.locator(cells ? '#perf-run' : '#perf-probe').click();
  interval = setInterval(() => {
    saved = saved
      .then(async () => {
        const rows = await page.locator('#perf-results').innerText();
        if (rows && rows !== previous) {
          previous = rows;
          process.stdout.write(rows.split('\n').at(-1) + '\n');
          await save();
        }
      })
      .catch((error) => process.stderr.write(`Progress snapshot: ${error.message}\n`));
  }, 10000);
  try {
    await page.locator('#perf-progress').filter({ hasText: 'Finished' }).waitFor({ timeout });
  } catch (error) {
    await page.locator('#perf-stop').click();
    await saved;
    await save();
    throw error;
  }
  clearInterval(interval);
  await saved;
  await memorySampler?.take();
  if (
    servedArtifacts &&
    (JSON.stringify(await verifyServedBuild()) !== JSON.stringify(servedArtifacts) ||
      JSON.stringify(localBuildFingerprint()) !== JSON.stringify(localBuildDigest))
  )
    throw new Error('Build changed during capacity run');
  const result = await save();
  process.stdout.write(
    JSON.stringify(
      {
        largestPassed: result.largestPassed,
        smallestFailed: result.smallestFailed,
        budgetMiB: result.budgetMiB,
        output,
      },
      null,
      2,
    ) + '\n',
  );
} finally {
  clearInterval(interval);
  await memorySampler?.stop();
  if (profiler) {
    try {
      const { profile } = await profiler.send('Profiler.stop');
      const path = resolve(cpuProfilePath);
      mkdirSync(dirname(path), { recursive: true });
      writeFileSync(path, JSON.stringify(profile));
    } finally {
      await browser.close();
    }
  } else await browser.close();
}
