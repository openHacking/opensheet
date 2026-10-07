export const LIMIT = 10_000_000;
export const STEP = 10_000;
export const thresholds = { editP95Ms: 100, scrollP95Ms: 33, restoreMs: 10_000 };
export const percentile = (values: number[], p = 0.95) =>
  [...values].sort((a, b) => a - b)[Math.max(0, Math.ceil(values.length * p) - 1)] ?? 0;
export interface Round {
  restoreMs: number;
  verifyMs: number;
  editP95Ms: number;
  scrollP95Ms: number;
  bytes: number;
  mainHeapBytes?: number;
  workerCacheBytes?: number;
  viewportCacheBytes?: number;
}
export interface Measurement {
  elapsedMs?: number;
  cells: number;
  passed: boolean;
  failure?: string;
  boundary?: string;
  rounds: Round[];
}
export async function probeCapacity(
  test: (count: number) => Promise<Measurement>,
  stopped: () => boolean,
  max = LIMIT,
) {
  const results: Measurement[] = [];
  let low = 0,
    high: number | undefined,
    count = Math.min(STEP, max);
  while (!stopped()) {
    const result = await test(count);
    results.push(result);
    if (result.failure === 'ABORTED') break;
    if (result.passed) low = Math.max(low, count);
    else high = high === undefined ? count : Math.min(high, count);
    if (result.failure === 'QuotaExceededError') break;
    if (result.passed && count === max) {
      result.boundary = 'Dimension boundary';
      break;
    }
    if (high !== undefined) {
      if (high - low <= STEP) break;
      count = Math.floor((low + high) / (2 * STEP)) * STEP;
      if (count <= low) break;
    } else count = Math.min(max, count * 2);
  }
  return { results, largestPassed: low, smallestFailed: high };
}
