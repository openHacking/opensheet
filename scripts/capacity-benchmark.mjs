import { chromium } from '@playwright/test';
import { readFileSync, writeFileSync, mkdirSync, renameSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { arch, platform } from 'node:os';
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
const browser = await chromium.launch();
let interval,
  saved = Promise.resolve(),
  previous = '';
try {
  const page = await browser.newPage({ viewport: { width: 1280, height: 800 } });
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
      engineFormat: 'paged-v4',
      browserVersion: browser.version(),
    });
    mkdirSync(dirname(output), { recursive: true });
    writeFileSync(output + '.tmp', JSON.stringify(result, null, 2) + '\n');
    renameSync(output + '.tmp', output);
    return result;
  };
  await page.locator('#perf-probe').click();
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
  await browser.close();
}
