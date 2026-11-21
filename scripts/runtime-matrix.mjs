import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { execFileSync, spawnSync } from 'node:child_process';
import {
  closeSync,
  existsSync,
  mkdtempSync,
  openSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from 'node:fs';
import { tmpdir, platform, release, arch } from 'node:os';
import { join } from 'node:path';

const output = process.argv[2];
if (!output) throw new Error('Usage: node scripts/runtime-matrix.mjs NEW_REPORT.json');
const tests = [
  'tests/e2e/order-workload.spec.ts',
  'tests/e2e/opaque-sandbox-capabilities.spec.ts',
  'tests/e2e/broadcast-channel-startup.spec.ts',
  'tests/e2e/worker-startup-denial.spec.ts',
  'tests/e2e/worker-csp-startup.spec.ts',
  'tests/e2e/malformed-indexeddb.spec.ts',
  'tests/e2e/malformed-web-locks.spec.ts',
  'tests/e2e/lock-handover.spec.ts',
];
const sources = [
  'packages/core/src/workbook.ts',
  'packages/core/src/engine.ts',
  'packages/core/src/storage.ts',
  'packages/core/src/engine.worker.js',
  'packages/core/src/worker-runtime.ts',
  'apps/playground/src/routing.ts',
  'playwright.config.ts',
  'pnpm-lock.yaml',
  'scripts/runtime-matrix.mjs',
  ...tests,
];
const sha256 = (path) => createHash('sha256').update(readFileSync(path)).digest('hex');
const git = (...args) => execFileSync('git', args, { encoding: 'utf8' }).trim();
if (existsSync(output)) throw new Error('Refusing to overwrite an existing report');
const before = Object.fromEntries(sources.map((path) => [path, sha256(path)]));
const headBefore = git('rev-parse', 'HEAD');
const version = {};
for (const [name, browserType] of Object.entries({ chromium, firefox, webkit })) {
  const browser = await browserType.launch({ headless: true });
  try {
    version[name] = browser.version();
  } finally {
    await browser.close();
  }
}
const temporary = mkdtempSync(join(tmpdir(), 'opensheet-runtime-matrix-'));
const reporterPath = join(temporary, 'report.json'),
  errorPath = join(temporary, 'stderr.txt');
const stdout = openSync(reporterPath, 'w'),
  stderr = openSync(errorPath, 'w');
try {
  const started = new Date().toISOString();
  const run = spawnSync(
    'pnpm',
    ['exec', 'playwright', 'test', ...tests, '--workers=3', '--reporter=json'],
    { stdio: ['ignore', stdout, stderr], encoding: 'utf8', env: { ...process.env, CI: '1' } },
  );
  closeSync(stdout);
  closeSync(stderr);
  const raw = readFileSync(reporterPath, 'utf8');
  let playwright;
  try {
    playwright = JSON.parse(raw);
  } catch {
    playwright = { parseError: 'Playwright did not produce JSON', raw };
  }
  const after = Object.fromEntries(sources.map((path) => [path, sha256(path)]));
  const consistent =
    JSON.stringify(before) === JSON.stringify(after) && headBefore === git('rev-parse', 'HEAD');
  const report = {
    started,
    finished: new Date().toISOString(),
    command: ['pnpm', 'exec', 'playwright', 'test', ...tests, '--workers=3', '--reporter=json'],
    freshServer: true,
    scope:
      'Local headless Playwright on one host; injected missing capabilities and CSP plus opaque/same-origin frames, task workflow, persistence and lock handover. Not native Safari, a real private-mode trial, a physical device matrix or a guarantee against future storage eviction.',
    environment: {
      platform: platform(),
      release: release(),
      arch: arch(),
      node: process.version,
      browserVersions: version,
    },
    source: {
      commit: headBefore,
      status: git('status', '--short'),
      sha256: before,
      sha256After: after,
      consistent,
    },
    exitCode: run.status,
    signal: run.signal,
    spawnError: run.error?.message ?? null,
    stderr: readFileSync(errorPath, 'utf8'),
    playwright,
  };
  writeFileSync(output, JSON.stringify(report, null, 2) + '\n', { flag: 'wx' });
  console.log(
    JSON.stringify({ output, exitCode: run.status, stats: playwright.stats, versions: version }),
  );
  if (
    run.status !== 0 ||
    run.error ||
    !consistent ||
    !playwright.stats?.expected ||
    playwright.stats.unexpected ||
    playwright.stats.skipped ||
    playwright.stats.flaky
  )
    process.exitCode = 1;
} finally {
  rmSync(temporary, { recursive: true, force: true });
}
