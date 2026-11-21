# History recovery on a million-cell workbook

Measured on 2026-11-21 with `scripts/history-large-paired.mjs` against the local production preview on port 5186. The script alternates three fresh headless Chromium contexts for the current Worker and an esbuild Worker using the engine source from `da9ade8`, before historical-directory and transition validation was added. The compared commits differ only in `packages/core/src/engine.ts` under `packages/`; both substituted Workers use the current surrounding sources and the same esbuild configuration. Their hashes, source/build fingerprints, browser version and environment are in the reports. The worktree also contained separately staged capacity artifacts and scripts; this is not a clean-release build measurement.

Each sample generates one million mixed values in 100 columns, applies 100 consecutive single-cell edits, reopens three times, undoes all 100 edits, reopens after the 50th undo, and redoes all 100 edits. It checks the revision and edited value on each reopen, checks the edited value after every undo/redo, and compares every generated cell after redo. Timings for undo and redo measure the API call only; verification reads and the midpoint reopen are outside those timings. Each run verified 1,000,000 cells.

| Measurement                                         |      Earlier engine |      Current engine |
| --------------------------------------------------- | ------------------: | ------------------: |
| Three initial reopens per round, median (range), ms | 150.5 (147.9–155.9) | 149.9 (148.3–168.1) |
| Reopen after 50 undos, three samples, ms            | 162.0, 155.5, 149.8 | 143.3, 154.8, 153.8 |
| Total time for 100 undos, three samples, ms         | 187.6, 167.5, 170.7 | 209.7, 196.5, 199.5 |
| Total time for 100 redos, three samples, ms         | 182.3, 162.4, 160.5 | 204.2, 169.2, 183.3 |

[Raw midpoint results](../benchmarks/history-large-paired-midpoint-2026-11-21.json) retain each reopen and history operation, paired order, all-cell counts, package/bundle hashes and storage statistics. The three paired undo totals were 22.1–29.0 ms higher for the current engine, while reopen ranges overlap. These are small local samples, not a measured causal breakdown or a latency percentile. The fixture repeatedly changes one page, has no styles, and does not cover large metadata, distinct-page edits, other browsers, devices, OS-cold caches or Worker/process peak memory. The two Workers are esbuild substitutes rather than the shipped Vite Worker, and host scheduling can affect both variants.

An earlier [pilot](../benchmarks/history-large-paired-2026-11-21.json) omitted the midpoint reopen; its first baseline history timings were substantially higher than its later rounds, so it should not be treated as a stable baseline. The [first failed attempt](../benchmarks/history-large-paired-2026-11-21-failed-read-limit.json) used 25,000-cell verification batches and hit the 16,384-cell read limit. The final script uses 10,000-cell batches. Both prior artifacts are retained; the pilot used the earlier version of the script without midpoint reopening.
