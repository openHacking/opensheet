export const LIMIT = 10_000_000;
export const STEP = 10_000;
/** Warmups use row zero; the 100 measured edits span both ends of the worksheet. */
export const capacityEditRow = (index: number, rows: number) =>
  Math.floor((Math.max(0, index) * (rows - 1)) / 99);
/** Warmups use the top; the 300 measured scrolls span the whole vertical extent. */
export const capacityScrollOffset = (index: number, maxOffset: number) =>
  Math.floor((Math.max(0, index) * maxOffset) / 299);
/** A permutation spreads 300 horizontal samples across both edges without coupling them to row order. */
export const capacityHorizontalScrollOffset = (index: number, maxOffset: number) =>
  Math.floor((((Math.max(0, index) * 137) % 300) * maxOffset) / 299);
export const thresholds = { editP95Ms: 100, scrollP95Ms: 33, restoreMs: 10_000 };
export const percentile = (values: number[], p = 0.95) =>
  [...values].sort((a, b) => a - b)[Math.max(0, Math.ceil(values.length * p) - 1)] ?? 0;
export interface Round {
  restoreMs: number;
  verifyMs: number;
  editP95Ms: number;
  scrollP95Ms: number;
  /** Recorded samples exclude protocol warmups; optional for historical artifacts. */
  editSamplesMs?: number[];
  commitSamplesMs?: number[];
  viewportSamplesMs?: number[];
  commitP95Ms?: number;
  viewportP95Ms?: number;
  workerQueueSamplesMs?: number[];
  workerExecutionSamplesMs?: number[];
  rpcSamplesMs?: number[];
  workerActiveSamples?: string[];
  workerQueuedAheadSamples?: number[];
  scrollReadTraces?: Array<{
    count: number;
    maxRoundTripMs: number;
    sumExecutionMs: number;
    maxExecutionMs: number;
    maxQueueMs: number;
    maxOutsideWorkerMs: number;
    sumOutsideWorkerMs: number;
    pageCacheHits: number;
    pageCacheMisses: number;
    absentPages: number;
    storageMs: number;
    decodeMs: number;
    inflateMs: number;
    pageDecodeMs: number;
  }>;
  scrollPhaseTraces?: Array<{
    layoutMs: number;
    previousFetchMs: number;
    viewportFetchMs: number;
    frameWaitMs: number;
    drawMs: number;
  }>;
  scrollSamplesMs?: number[];
  viewportChecks?: number;
  scrollExtents?: {
    maxTop: number;
    maxLeft: number;
    observedTop: [number, number];
    observedLeft: [number, number];
  };
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
