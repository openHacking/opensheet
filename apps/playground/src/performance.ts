import {
  Workbook,
  createWorkbookFile,
  fixtureValue,
  type OpenSheet,
  type OperationTiming,
  type WorkbookOptions,
} from 'opensheet';
import { escapeHTML } from '@opensheetjs/formats';
import { $ } from './dom.js';
import {
  LIMIT,
  capacityEditRow,
  capacityHorizontalScrollOffset,
  capacityScrollOffset,
  percentile,
  probeCapacity,
  thresholds,
  type Measurement,
  type Round,
} from './performance-model.js';
import type { PlaygroundState } from './state.js';
const ID = 'performance-workbook';
const DATABASE = 'opensheet-performance';
const mib = (n: number) => (n / 1048576).toFixed(2);
const frame = () => new Promise<number>((resolve) => requestAnimationFrame(resolve));
export function createPerformance(
  app: OpenSheet,
  state: PlaygroundState,
  download: (data: BlobPart, name: string, type?: string) => void,
) {
  let task: Promise<void> | undefined;
  let mode: 'single' | 'capacity' = 'single';
  let running = false,
    stopped = false,
    abort = new AbortController(),
    progress = 'Ready',
    budgetMiB = 256,
    target = 10000;
  let results: Measurement[] = [],
    quota: { usage?: number; quota?: number } = {},
    lastBook: Workbook | undefined;
  const traceOperations = new URLSearchParams(location.search).has('trace-ops');
  let lastExecuteTiming: OperationTiming | undefined;
  let currentScrollReads: OperationTiming[] | undefined;
  const getLastExecuteTiming = () => lastExecuteTiming;
  let largest = 0,
    smallest: number | undefined;
  const options = (): WorkbookOptions => ({
    database: DATABASE,
    preserveOrphans: true,
    ...(traceOperations
      ? {
          onOperationTiming: (timing: OperationTiming) => {
            if (timing.method === 'execute') lastExecuteTiming = timing;
            if (timing.method === 'read') currentScrollReads?.push(timing);
          },
        }
      : {}),
    budgetBytes: budgetMiB * 1048576,
    onProgress: (p) => {
      progress = `${p.phase}: ${p.completed.toLocaleString()}${p.total ? ' / ' + p.total.toLocaleString() : ''}`;
      update();
    },
  });
  async function estimate() {
    try {
      quota = (await navigator.storage?.estimate()) ?? {};
    } catch {
      quota = {};
    }
  }
  const shell = () =>
    `<div class="performance-controls"><label>Cells<input id="perf-cells" type="number" min="10000" max="${LIMIT}" step="10000" value="${target}"></label><label>Budget (MiB)<input id="perf-budget" type="number" min="1" max="8192" value="${budgetMiB}"></label><div class="performance-buttons"><button class="button primary" id="perf-run">Test size</button><button class="button secondary" id="perf-probe">Find capacity</button><button class="button secondary" id="perf-stop">Stop</button><button class="button secondary" id="perf-restore">Restore</button><button class="button secondary" id="perf-clear">Clear test data</button><button class="button secondary" id="perf-download">Download results</button></div></div><p>100 columns · 60% numbers / 30% 32-byte text / 10% booleans · fixed seed. Each size: 3 rounds, 100 edits, 300 scrolling frames.</p><p>Pass: edit P95 ≤100 ms · complete scroll frame P95 ≤33 ms · reopen to editable viewport ≤10 s. Formula and XLSX performance are separate workloads.</p><p id="perf-progress" role="status" aria-live="polite"></p><div id="perf-summary"></div><div class="performance-results"><table><thead><tr><th>Cells</th><th>Result</th><th>Edit P95</th><th>Scroll P95</th><th>Restore</th><th>Verify</th><th>Elapsed</th><th>MiB</th><th>Bytes/cell</th><th>Worker cache</th><th>Viewport cache</th><th>Main heap</th></tr></thead><tbody id="perf-results"></tbody></table></div>`;
  function update() {
    if (state.currentScene !== 'performance' || !$('perf-progress')) return;
    app.element.inert = running;
    $('perf-progress').textContent = progress;
    $('perf-summary').textContent =
      `Largest passed: ${largest.toLocaleString()} · Smallest failed: ${smallest?.toLocaleString() ?? '—'} · Origin usage: ${quota.usage === undefined ? 'unknown' : mib(quota.usage) + ' MiB'} · Estimated origin quota: ${quota.quota === undefined ? 'unknown' : mib(quota.quota) + ' MiB'} (includes other same-origin storage). Application bytes count encoded records, staging and history; browser disk usage is a separate estimate.`;
    $('perf-results').innerHTML = results
      .map((r) => {
        const worst = (
          k: keyof Omit<
            Round,
            | 'editSamplesMs'
            | 'scrollSamplesMs'
            | 'commitSamplesMs'
            | 'viewportSamplesMs'
            | 'workerQueueSamplesMs'
            | 'workerExecutionSamplesMs'
            | 'rpcSamplesMs'
            | 'workerActiveSamples'
            | 'workerQueuedAheadSamples'
            | 'scrollReadTraces'
            | 'scrollPhaseTraces'
            | 'scrollExtents'
          >,
        ) => Math.max(0, ...r.rounds.map((v) => v[k] ?? 0));
        return `<tr><td>${r.cells.toLocaleString()}</td><td>${escapeHTML(r.passed ? (r.boundary ?? 'Passed') : r.failure === 'STORAGE_BUDGET' ? 'Test boundary: application budget' : (r.failure ?? 'Failed'))}</td><td>${worst('editP95Ms').toFixed(1)} ms</td><td>${worst('scrollP95Ms').toFixed(1)} ms</td><td>${worst('restoreMs').toFixed(0)} ms</td><td>${worst('verifyMs').toFixed(0)} ms</td><td>${(r.elapsedMs ?? 0).toFixed(0)} ms</td><td>${mib(worst('bytes'))}</td><td>${(worst('bytes') / r.cells).toFixed(2)}</td><td>${mib(worst('workerCacheBytes'))} MiB</td><td>${mib(worst('viewportCacheBytes'))} MiB</td><td>${r.rounds.some((v) => v.mainHeapBytes !== undefined) ? mib(worst('mainHeapBytes')) + ' MiB' : 'unknown'}</td></tr>`;
      })
      .join('');
    for (const id of [
      'perf-run',
      'perf-probe',
      'perf-restore',
      'perf-clear',
      'perf-cells',
      'perf-budget',
    ])
      $<HTMLButtonElement>(id).disabled = running;
    $<HTMLButtonElement>('perf-stop').disabled = !running;
    $<HTMLSelectElement>('scene-select').disabled = running;
    for (const el of document.querySelectorAll<
      HTMLButtonElement | HTMLInputElement | HTMLSelectElement
    >(
      '.document-actions button, .utility-bar input, .utility-bar button, .utility-bar select, #reset-demo',
    ))
      el.disabled = running;
  }
  async function empty() {
    app.getGrid()?.dispose();
    if (lastBook) {
      try {
        await lastBook.deleteStorage();
        await lastBook.close();
      } catch {
        lastBook.dispose();
        // A failed/closed Worker cannot delete its database. Reopen the durable head first.
        try {
          const stale = await app.open(ID, options());
          await stale.deleteStorage();
          await stale.close();
        } catch (error) {
          if ((error as { code?: string }).code !== 'INVALID_ARGUMENT') throw error;
        }
      }
    } else {
      try {
        const book = await app.open(ID, options());
        await book.deleteStorage();
        await book.close();
      } catch {
        /* Database may not exist yet. */
      }
    }
    return createEmpty();
  }
  async function createEmpty() {
    const snapshot = createWorkbookFile({
      sheets: [{ name: 'Performance', rows: 100, columns: 100 }],
    });
    snapshot.workbookId = ID;
    lastBook = await new Workbook(snapshot, options()).ready();
    await app.attachWorkbook(lastBook);
    return lastBook;
  }
  async function start() {
    state.currentScene = 'performance';
    $('scene-insights').innerHTML = shell();
    bind();
    await estimate();
    try {
      lastBook = await app.open(ID, options());
    } catch (error) {
      if ((error as { code?: string }).code !== 'INVALID_ARGUMENT') throw error;
      await createEmpty();
    }
    state.dirty = false;
    progress = 'Ready — measurements depend on this browser and device.';
    update();
  }
  async function verify(book: Workbook, count: number) {
    const sheet = book.getSheets()[0],
      sh = book.sheetData(sheet.id),
      start = performance.now();
    let seen = 0;
    const rows = new Map(sh.rowOrder.map((id, i) => [id, i])),
      cols = new Map(sh.columnOrder.map((id, i) => [id, i]));
    for await (const batch of book.streamRange(
      sheet.id,
      { startRow: 0, endRow: sh.rowOrder.length, startColumn: 0, endColumn: 100 },
      { signal: abort.signal, values: false },
    )) {
      for (const c of Object.values(batch.cells)) {
        const n = rows.get(c.rowId)! * 100 + cols.get(c.columnId)!;
        if (n >= count || !('value' in c.input) || c.input.value !== fixtureValue(n))
          throw new Error(`Data validation failed at cell ${n}`);
        seen++;
      }
      progress = `Verify: ${seen.toLocaleString()} / ${count.toLocaleString()}`;
      update();
    }
    if (seen !== count) throw new Error(`Expected ${count} populated cells, recovered ${seen}`);
    return performance.now() - start;
  }
  async function testSize(count: number): Promise<Measurement> {
    const candidateStart = performance.now();
    const result: Measurement = { cells: count, passed: false, rounds: [] };
    try {
      const book = await empty();
      abort.signal.throwIfAborted();
      await book.generate(count, budgetMiB * 1048576);
      await app.attachWorkbook(book);
      app.getGrid()!.setReadOnly(true);
      for (let round = 0; round < 3; round++) {
        abort.signal.throwIfAborted();
        progress = `Round ${round + 1}/3 · reopen ${count.toLocaleString()} cells`;
        update();
        const restoreStart = performance.now();
        lastBook = await app.open(ID, options());
        const restoreMs = performance.now() - restoreStart;
        const current = lastBook;
        app.getGrid()!.setReadOnly(true);
        const verifyMs = await verify(current, count),
          sheet = current.getSheets()[0];
        const edits: number[] = [],
          commits: number[] = [],
          viewports: number[] = [],
          workerQueues: number[] = [],
          workerExecutions: number[] = [],
          rpcs: number[] = [],
          workerActive: string[] = [],
          workerAhead: number[] = [];
        for (let i = -5; i < 100; i++) {
          abort.signal.throwIfAborted();
          const row = capacityEditRow(i, sheet.rowCount),
            column = (Math.max(0, i) * 37) % 100;
          app.selection.set({
            sheetId: sheet.id,
            startRow: row,
            endRow: row + 1,
            startColumn: column,
            endColumn: column + 1,
          });
          await app.getGrid()!.ready();
          const before = fixtureValue(row * 100 + column),
            start = performance.now();
          lastExecuteTiming = undefined;
          await sheet
            .range({ startRow: row, endRow: row + 1, startColumn: column, endColumn: column + 1 })
            .setValues([[i + 0.25]]);
          const committed = performance.now();
          await app.getGrid()!.ready();
          const ready = performance.now();
          if (i >= 0) {
            commits.push(committed - start);
            viewports.push(ready - committed);
            edits.push(ready - start);
            if (traceOperations) {
              const timing = getLastExecuteTiming();
              if (!timing) throw new Error('Worker timing was not reported');
              workerQueues.push(timing.queueMs);
              workerExecutions.push(timing.executionMs);
              rpcs.push(timing.roundTripMs);
              workerActive.push(timing.activeAtArrival);
              workerAhead.push(timing.queuedAhead);
            }
          }
          if (!(await current.undo())) throw new Error('Undo was unavailable');
          if (
            (
              await sheet
                .range({
                  startRow: row,
                  endRow: row + 1,
                  startColumn: column,
                  endColumn: column + 1,
                })
                .getValues()
            )[0][0] !== before
          )
            throw new Error('Undo validation failed');
          if (!(await current.redo())) throw new Error('Redo was unavailable');
          if (
            (
              await sheet
                .range({
                  startRow: row,
                  endRow: row + 1,
                  startColumn: column,
                  endColumn: column + 1,
                })
                .getValues()
            )[0][0] !==
            i + 0.25
          )
            throw new Error('Redo validation failed');
          if (!(await current.undo())) throw new Error('Final undo was unavailable');
        }
        const scroller = app.getGrid()!.scroller,
          frames: number[] = [],
          scrollReadTraces: NonNullable<Round['scrollReadTraces']> = [],
          scrollPhaseTraces: NonNullable<Round['scrollPhaseTraces']> = [];
        let viewportChecks = 0;
        const scrollExtents = {
          maxTop: Math.max(0, scroller.scrollHeight - scroller.clientHeight),
          maxLeft: Math.max(0, scroller.scrollWidth - scroller.clientWidth),
          observedTop: [Infinity, -Infinity] as [number, number],
          observedLeft: [Infinity, -Infinity] as [number, number],
        };
        const geometry = current.sheetData(sheet.id);
        if (
          Object.keys(geometry.rows).length ||
          Object.keys(geometry.columns).length ||
          geometry.freeze.rows ||
          geometry.freeze.columns
        )
          throw new Error('Unexpected performance fixture geometry');
        for (let i = -30; i < 300; i++) {
          abort.signal.throwIfAborted();
          const start = await frame();
          const reads: OperationTiming[] | undefined = traceOperations ? [] : undefined;
          const phases = traceOperations
            ? {
                layoutMs: -1,
                previousFetchMs: -1,
                viewportFetchMs: -1,
                frameWaitMs: -1,
                drawMs: -1,
              }
            : undefined;
          currentScrollReads = reads;
          try {
            scroller.scrollTop = capacityScrollOffset(
              i,
              Math.max(0, scroller.scrollHeight - scroller.clientHeight),
            );
            scroller.scrollLeft = capacityHorizontalScrollOffset(
              i,
              Math.max(0, scroller.scrollWidth - scroller.clientWidth),
            );
            await app.getGrid()!.ready(phases);
          } finally {
            currentScrollReads = undefined;
          }
          const end = performance.now();
          if (i >= 0) {
            frames.push(end - start);
            scrollExtents.observedTop[0] = Math.min(
              scrollExtents.observedTop[0],
              scroller.scrollTop,
            );
            scrollExtents.observedTop[1] = Math.max(
              scrollExtents.observedTop[1],
              scroller.scrollTop,
            );
            scrollExtents.observedLeft[0] = Math.min(
              scrollExtents.observedLeft[0],
              scroller.scrollLeft,
            );
            scrollExtents.observedLeft[1] = Math.max(
              scrollExtents.observedLeft[1],
              scroller.scrollLeft,
            );
            const zoom = app.getGrid()!.getZoom();
            const rows = [scroller.scrollTop, scroller.scrollTop + scroller.clientHeight].map(
              (offset) => Math.min(sheet.rowCount - 1, Math.floor(offset / (30 * zoom))),
            );
            const columns = [scroller.scrollLeft, scroller.scrollLeft + scroller.clientWidth].map(
              (offset) => Math.min(sheet.columnCount - 1, Math.floor(offset / (128 * zoom))),
            );
            for (let row = rows[0]; row <= rows[1]; row++)
              for (let column = columns[0]; column <= columns[1]; column++) {
                const value = fixtureValue(row * 100 + column);
                const display =
                  typeof value === 'boolean' ? String(value).toUpperCase() : String(value);
                if (
                  !current.isLoaded(sheet.id, row, column) ||
                  current.peekValue(sheet.id, row, column) !== value ||
                  current.peekDisplay(sheet.id, row, column) !== display
                )
                  throw new Error(`Viewport validation failed at row ${row}, column ${column}`);
                viewportChecks++;
              }
            if (reads)
              scrollReadTraces.push({
                count: reads.length,
                maxRoundTripMs: Math.max(0, ...reads.map((r) => r.roundTripMs)),
                sumExecutionMs: reads.reduce((sum, r) => sum + r.executionMs, 0),
                maxExecutionMs: Math.max(0, ...reads.map((r) => r.executionMs)),
                maxQueueMs: Math.max(0, ...reads.map((r) => r.queueMs)),
                maxOutsideWorkerMs: Math.max(
                  0,
                  ...reads.map((r) => Math.max(0, r.roundTripMs - r.queueMs - r.executionMs)),
                ),
                sumOutsideWorkerMs: reads.reduce(
                  (sum, r) => sum + Math.max(0, r.roundTripMs - r.queueMs - r.executionMs),
                  0,
                ),
                pageCacheHits: reads.reduce((sum, r) => sum + (r.read?.cacheHits ?? 0), 0),
                pageCacheMisses: reads.reduce((sum, r) => sum + (r.read?.cacheMisses ?? 0), 0),
                absentPages: reads.reduce((sum, r) => sum + (r.read?.absentPages ?? 0), 0),
                storageMs: reads.reduce((sum, r) => sum + (r.read?.storageMs ?? 0), 0),
                decodeMs: reads.reduce((sum, r) => sum + (r.read?.decodeMs ?? 0), 0),
                inflateMs: reads.reduce((sum, r) => sum + (r.read?.inflateMs ?? 0), 0),
                pageDecodeMs: reads.reduce((sum, r) => sum + (r.read?.pageDecodeMs ?? 0), 0),
              });
            if (phases) scrollPhaseTraces.push(phases);
          }
        }
        if (
          scrollExtents.observedTop[0] > 1 ||
          scrollExtents.observedLeft[0] > 1 ||
          scrollExtents.observedTop[1] < scrollExtents.maxTop - 1 ||
          scrollExtents.observedLeft[1] < scrollExtents.maxLeft - 1 ||
          Math.abs(
            Math.max(0, scroller.scrollHeight - scroller.clientHeight) - scrollExtents.maxTop,
          ) > 1 ||
          Math.abs(
            Math.max(0, scroller.scrollWidth - scroller.clientWidth) - scrollExtents.maxLeft,
          ) > 1
        )
          throw new Error('Scroll extent validation failed');
        await app.getGrid()!.ready();
        const mainHeapBytes = (performance as Performance & { memory?: { usedJSHeapSize: number } })
          .memory?.usedJSHeapSize;
        const data: Round = {
          restoreMs,
          verifyMs,
          editP95Ms: percentile(edits),
          scrollP95Ms: percentile(frames),
          editSamplesMs: edits,
          commitSamplesMs: commits,
          viewportSamplesMs: viewports,
          commitP95Ms: percentile(commits),
          viewportP95Ms: percentile(viewports),
          ...(traceOperations
            ? {
                workerQueueSamplesMs: workerQueues,
                workerExecutionSamplesMs: workerExecutions,
                rpcSamplesMs: rpcs,
                workerActiveSamples: workerActive,
                workerQueuedAheadSamples: workerAhead,
                scrollReadTraces,
                scrollPhaseTraces,
              }
            : {}),
          scrollSamplesMs: frames,
          viewportChecks,
          scrollExtents,
          bytes: current.stats.bytes,
          mainHeapBytes,
          workerCacheBytes: (await current.storageStats()).cacheBytes,
          viewportCacheBytes: current.stats.viewportCacheBytes,
        };
        result.rounds.push(data);
        if (
          restoreMs > thresholds.restoreMs ||
          data.editP95Ms > thresholds.editP95Ms ||
          data.scrollP95Ms > thresholds.scrollP95Ms
        )
          throw new Error('Interaction threshold exceeded');
      }
      result.passed = true;
    } catch (error) {
      result.failure = stopped
        ? 'ABORTED'
        : ((error as { code?: string }).code ??
          (error instanceof Error ? error.message : String(error)));
    }
    result.elapsedMs = performance.now() - candidateStart;
    results.push(result);
    if (result.passed) largest = Math.max(largest, count);
    else if (result.failure !== 'ABORTED') smallest = Math.min(smallest ?? Infinity, count);
    await estimate();
    update();
    return result;
  }
  async function run(probe: boolean) {
    if (running) return;
    target = Number($<HTMLInputElement>('perf-cells').value);
    budgetMiB = Number($<HTMLInputElement>('perf-budget').value);
    if (
      !Number.isSafeInteger(target) ||
      target < 10000 ||
      target > LIMIT ||
      target % 10000 !== 0 ||
      !Number.isFinite(budgetMiB) ||
      budgetMiB < 1 ||
      budgetMiB > 8192
    ) {
      progress = 'Enter a valid cell count and budget.';
      update();
      return;
    }
    await estimate();
    if (
      quota.quota !== undefined &&
      quota.usage !== undefined &&
      budgetMiB * 1048576 > Math.max(0, quota.quota - quota.usage) * 0.8
    ) {
      progress = 'Budget exceeds 80% of estimated remaining origin quota. Choose a smaller budget.';
      update();
      return;
    }
    mode = probe ? 'capacity' : 'single';
    running = true;
    stopped = false;
    abort = new AbortController();
    results = [];
    largest = 0;
    smallest = undefined;
    update();
    try {
      if (probe) await probeCapacity(testSize, () => stopped);
      else await testSize(target);
    } finally {
      running = false;
      progress = stopped
        ? 'Stopped — incomplete tests do not count as passed.'
        : 'Finished — largest passed is an observed result, not a universal limit.';
      app.getGrid()?.setReadOnly(false);
      state.dirty = false;
      update();
    }
  }
  function bind() {
    $('perf-run').onclick = () => {
      task = run(false).catch((error) => {
        progress = String(error);
        update();
      });
    };
    $('perf-probe').onclick = () => {
      task = run(true).catch((error) => {
        progress = String(error);
        update();
      });
    };
    $('perf-stop').onclick = () => {
      stopped = true;
      abort.abort();
      lastBook?.cancel();
    };
    $('perf-restore').onclick = () => {
      void app
        .open(ID, options())
        .then((book) => {
          lastBook = book;
          progress = 'Restored to an editable viewport.';
          update();
        })
        .catch((error) => {
          progress = String(error);
          update();
        });
    };
    $('perf-clear').onclick = () => {
      void empty()
        .then(() => {
          progress = 'Test data cleared.';
          update();
        })
        .catch((error) => {
          progress = String(error);
          update();
        });
    };
    $('perf-download').onclick = () =>
      download(
        JSON.stringify(
          {
            schemaVersion: 1,
            engineFormat: 'paged',
            mode,
            status: running
              ? 'running'
              : stopped
                ? 'stopped'
                : results.length
                  ? 'finished'
                  : 'ready',
            protocol: {
              operationTiming: traceOperations,
              rounds: 3,
              editWarmups: 5,
              editSamples: 100,
              editRowSampling: 'evenly-spaced-inclusive-v1',
              scrollWarmups: 30,
              scrollSamples: 300,
              scrollSampling: 'evenly-spaced-2d-permuted-inclusive-v2',
              viewportValidation: 'viewport-rectangle-cached-value-display-v2',
              blockRows: 64,
              blockColumns: 32,
              workerCacheBudgetBytes: 64 * 1048576,
              viewportCacheBudgetBytes: 8 * 1048576,
            },
            date: new Date().toISOString(),
            browser: navigator.userAgent,
            hardwareConcurrency: navigator.hardwareConcurrency,
            deviceMemoryGiB: (navigator as Navigator & { deviceMemory?: number }).deviceMemory,
            viewport: { width: innerWidth, height: innerHeight },
            workload: {
              columns: 100,
              seed: 'LCG-1664525-1013904223',
              numbers: 0.6,
              text: 0.3,
              textBytes: 32,
              booleans: 0.1,
            },
            budgetMiB,
            thresholds,
            quota,
            memoryScope:
              'mainHeapBytes covers the main document where supported; workerCacheBytes and viewportCacheBytes are conservative encoded-byte reservations, not total process heap',
            largestPassed: largest,
            smallestFailed: smallest,
            results,
          },
          null,
          2,
        ),
        'opensheet-performance.json',
        'application/json',
      );
  }
  return {
    start,
    update,
    leave: async () => {
      stopped = true;
      abort.abort();
      lastBook?.cancel();
      await task;
    },
    stop: () => {
      stopped = true;
      abort.abort();
      lastBook?.cancel();
    },
    dispose: () => {
      stopped = true;
      abort.abort();
    },
  };
}
