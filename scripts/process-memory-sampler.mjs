import { execFile } from 'node:child_process';
import { platform } from 'node:os';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

// RSS is summed per process and can count shared pages more than once.
export function summarizeProcessTable(table, parentPid) {
  const processes = table.split('\n').flatMap((line) => {
    const match = line.match(/^\s*(\d+)\s+(\d+)\s+(\d+)\s+(.+)$/);
    return match
      ? [
          {
            pid: Number(match[1]),
            ppid: Number(match[2]),
            rss: Number(match[3]) * 1024,
            command: match[4],
          },
        ]
      : [];
  });
  const descendants = new Set([parentPid]);
  let changed = true;
  while (changed) {
    changed = false;
    for (const item of processes) {
      if (descendants.has(item.ppid) && !descendants.has(item.pid)) {
        descendants.add(item.pid);
        changed = true;
      }
    }
  }
  const totals = {};
  for (const item of processes) {
    if (
      !descendants.has(item.pid) ||
      !/chrome-headless-shell|chromium_headless_shell|\/Chromium(?:\.app|\/)|\/Google Chrome for Testing/i.test(
        item.command,
      )
    )
      continue;
    const type = item.command.match(/--type=([^\s]+)/)?.[1];
    const role =
      type === 'renderer'
        ? 'renderer'
        : type === 'gpu-process'
          ? 'gpu'
          : type === 'utility'
            ? 'utility'
            : type
              ? 'other'
              : 'browser';
    totals[role] ??= { count: 0, rssBytes: 0 };
    totals[role].count++;
    totals[role].rssBytes += item.rss;
  }
  return {
    roles: totals,
    totalRssBytes: Object.values(totals).reduce((sum, role) => sum + role.rssBytes, 0),
  };
}

export function startProcessMemorySampler({ intervalMs = 1000, parentPid = process.pid } = {}) {
  if (!['darwin', 'linux'].includes(platform()))
    throw new Error('Process memory sampling requires macOS or Linux');
  const started = performance.now();
  const samples = [];
  const errors = [];
  let pending;
  async function take() {
    if (pending) return pending;
    pending = (async () => {
      try {
        const { stdout } = await execFileAsync('ps', ['-axo', 'pid=,ppid=,rss=,command='], {
          maxBuffer: 16 * 1024 * 1024,
        });
        samples.push({
          elapsedMs: Math.round(performance.now() - started),
          ...summarizeProcessTable(stdout, parentPid),
        });
      } catch (error) {
        errors.push(String(error.message));
      }
    })();
    try {
      await pending;
    } finally {
      pending = undefined;
    }
  }
  const timer = setInterval(() => void take(), intervalMs);
  return {
    take,
    snapshot: () => ({
      intervalMs,
      rssSemantics:
        'Sum of browser process RSS; shared pages may be counted multiple times. Renderer processes may contain both page and Worker; separate attribution is unavailable.',
      samples: [...samples],
      errors: [...errors],
    }),
    stop: async () => {
      clearInterval(timer);
      await take();
    },
  };
}
