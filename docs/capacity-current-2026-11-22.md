# Local million-cell capacity run, 2026-11-22

A single local headless Chromium run of the current production build completed the existing three-round, one-million-cell protocol. The [full report](../benchmarks/capacity-0c457cb-2026-11-22-1m.json) retains all edit and scroll samples, viewport checks, build hashes and process samples. This is a measurement on one machine, not release acceptance or evidence that earlier failures are resolved.

| Round | Edit P95 | Scroll P95 | Reopen | Full value verification |
| --- | ---: | ---: | ---: | ---: |
| 1 | 33.4 ms | 19.2 ms | 132.4 ms | 1,971.5 ms |
| 2 | 33.4 ms | 19.2 ms | 130.7 ms | 1,966.5 ms |
| 3 | 33.3 ms | 19.3 ms | 133.0 ms | 1,973.7 ms |

The unchanged limits are 100 ms edit P95, 33 ms scroll P95 and 10 s reopen. Each round checked all one million generated values, 100 dispersed edits with undo/redo, and 28,009 cached viewport values/displays. The commit P95 was 12.1–12.6 ms and the subsequent viewport-ready P95 was 22.3–22.4 ms; these samples describe this protocol, not native input latency.

The Chromium process sampler collected 55 approximately one-second samples. Its highest sampled sum was 891.8 MiB RSS, including up to 654.1 MiB in the renderer process. This sum can double-count shared pages, misses peaks between samples, and cannot separate the Worker from the page. Reported Worker and viewport cache counters are reservations, not total heap. Chromium's main-document heap field was 17,100,000 bytes at each round boundary and does not measure the Worker. No memory acceptance threshold was inferred from these figures.

The source commit was `0c457cb991d6ededaafe1d075acf0abe34e171f4`; the worktree was dirty due to pre-existing staged evidence and documentation. The source fingerprint and served asset hashes are recorded in the report. The build was produced by the preceding `pnpm check`; the harness verified all served artifacts against that build before and after the run. The machine was an Apple M4 Mac mini with 16 GiB memory, and `pmset -g ps` reported AC power immediately before the run. The report records Chromium 153.0.8010.12, OS and hardware; other concurrent processes were not controlled.

Reproduce on a fresh checkout with `pnpm check`, start `pnpm preview --host 127.0.0.1 --port 4173 --strictPort`, then run `node scripts/capacity-benchmark.mjs --url=http://127.0.0.1:4173/ --cells=1000000 --verify-build --process-memory --power-mode=ac --output=benchmarks/NEW-NAME.json`. Check the power state independently and stop other workloads before measuring. The harness writes to its chosen output, so choose a new name to preserve earlier evidence.

The prior [failed and passing runs](performance.md) remain relevant. This run does not explain their variation, measure the Windows reference machine, exercise another browser or prove physical-device memory and interaction targets.
