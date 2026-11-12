# Generated workbook reference model

`tests/workbook-model.test.ts` compares the real Worker dispatcher and IndexedDB engine with an independent small in-memory model after each executed operation. It does not call product command reducers, formula parsers/evaluators, sort comparators or filter predicates to compute expectations.

The model tracks cell contents/results, stable row/column identities, merge rectangles, typed filter conditions and visible rows, revision, undo and redo states. Its formula language is deliberately small: a single absolute/relative cell reference multiplied by two, including recursive references, cycles and deleted-reference errors. This allows reference movement to be checked independently rather than duplicating a full spreadsheet implementation.

Generated operations include scalar/formula writes, row/column insertion/deletion, two-key sorting, all/any filters, clear filter, merge/unmerge, undo/redo, reopening with zero or small cache, rejected transactions and stale envelopes. Sorting freezes reference-model key values before movement; formula references move afterward. Structural edits transform references/conditions/header prefixes independently. Each trial deletes only its own generated workbook storage after all reopen/history assertions, then closes its Worker. This avoids accumulating unrelated test databases across hundreds of independent trials. The default run uses three fixed seeds, 12 sequences per seed and 40–70 candidate operations per sequence; dimensions remain within 1–8 rows/columns. Each seed must execute all 16 operation types. Invalid merge intersections and covered-cell writes are checked as rejections, while dimension-growth limits are preconditions for this small model.

Run the default suite with `pnpm test:model`. For an extended run, use `OPENSHEET_MODEL_SEED=20261023 OPENSHEET_MODEL_RUNS=100 pnpm test:model` and repeat in fresh processes for seeds `73013` and `991`. A single process with 100 sequences for all three seeds has twice reached the third test's 60-second limit; workbook storage cleanup did not eliminate that behavior. Isolating seeds preserves the full generated coverage; it does not resolve or explain the shared-process slowdown. To replay a fast-check failure, preserve the printed seed, path and counterexample, then run:

```sh
OPENSHEET_MODEL_SEED=20261023 OPENSHEET_MODEL_PATH='PATH_FROM_FAILURE' pnpm test:model
```

Set `OPENSHEET_MODEL_REPORT` to a new JSONL path to record per-seed operation counts; existing files are rejected. `OPENSHEET_MODEL_TIMEOUT_MS` can set a diagnostic test-run limit without changing product operation timeouts or performance gates.

When a run used a different sequence count, also set `OPENSHEET_MODEL_RUNS` to the recorded value. Fast-check shrinks failing sequences; confirmed product failures should become fixed regression fixtures alongside their seed/path. Replay mode skips aggregate operation-coverage assertions because a minimized trace naturally uses fewer operation kinds.

This supplements existing full formula corpora, resource-fault tests and browser tests. It does not establish arbitrary Excel formula correctness, maximum-dimension acceptance, clipboard/formatting behavior, multi-sheet reference coverage, retained-memory bounds or native accessibility. Those remain separate acceptance work.

## Test runner resource envelope

The shared Vitest configuration caps concurrent file workers at four, including ordinary unit and model tests. This bounds simultaneous in-memory storage fixtures rather than increasing test timeouts. The default 5-second ordinary-test threshold and explicit model-test limits remain unchanged. CLI overrides can be used for controlled diagnostics, but results should record the worker count; concurrent suite wall-clock timings are not product latency benchmarks.
