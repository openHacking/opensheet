# Revision-bound formula sorting design

Status: initial implementation and regression verification in progress. Tracks CALC-06 and DATA-03 multi-key sorting; full acceptance remains open.

## Problem and invariant

Before this change, `Engine.sort` rejected formula sort keys. It already uses pinned source pages, bounded external merge runs, stable original-row tie breaking, staged writes and atomic history publication. Its previous scalar comparator mixed numeric comparisons with string conversion and did not define a coherent ordering across mixed types; error inputs also became blank keys. Adding formula results without fixing these properties would produce misleading or revision-dependent order.

Every sort key must be computed from the same immutable pre-sort command-stage view. A transaction may edit cells and then sort before committing: the source is the stage immediately before that sort, not the public workbook head and not the progressively rearranged output. Cross-sheet dependencies use that same stage.

## Implementation boundary

Create a dedicated AsyncCalculation instance for the operation, with cell reads bound to the pinned `source.blocks` mapping and axis/sheet/date metadata at sort start. Never reuse `this.calculator`, whose reads target the committed head. Missing sheet/coordinate references must return REF errors. Do not trust imported caches for sorting; compute supported formulas from source inputs.

Calculate and store each key while generating initial runs, before any output page is changed. Merge passes compare persisted key values and never recompute formulas. Output movement preserves existing row identities, copies cells into destinations, offsets relative formula references exactly once, and removes external caches. Formulas can have different post-sort results if their moved relative references point outside the sorted range; the order is defined by pre-sort values, not a promise that recomputed values remain sorted.

## Ordering contract

Use a total type ordering: numbers, text, logical values, ordinary errors, then blanks. Blank keys stay last in either direction. Numbers compare numerically; text compares case-insensitively with deterministic UTF-16 lexicographic ordering; logical values order FALSE before TRUE; errors order by error code. Direction reverses nonblank ordering. Equal keys retain original source-row order, including case-insensitive ties. Empty strings are text, not blanks. Resource-limit errors must abort the operation rather than masquerade as ordinary sortable values.

This deliberately replaces the current mixed-type comparator. Document the compatibility change and verify scalar sorts as well as formula sorts. Multi-key sorting extends this ordering lexicographically over 1–64 distinct columns, each with its own direction. Original-row ordering breaks ties only after every key compares equal. The command accepts either the legacy `column`/`direction` pair or a `keys` array; ambiguous combinations are rejected. All key values are persisted in the initial runs and included in staging accounting.

## Resource and failure behavior

Keep a bounded calculation cache (at most 4 MiB) and deduct it from the existing 32 MiB operation working-set allowance. Include serialized computed keys in run-memory and scratch-storage accounting. Preserve cooperative yields/cancellation during reads, calculation, merge passes and output writes. Initial reads use scalar page access. Output combines modifications within a physical row band and flushes on band changes or before its accounted bytes plus the current row exceed the working budget. Pairwise merge uses at most 64 rows per buffer, scaled down by the largest persisted row; six buffer shares reserve room for readers, output and cursor/encoding transients. Large rows use the original two-row path with an explicit combined-size check. Serialized-size multipliers are conservative accounting estimates, not total process-heap measurements. Scratch and partially allocated pages remain unpublished until one successful revision commit; failure/cancellation must leave the previous workbook and history intact.

## Acceptance evidence required

- Hand-authored fixtures for numeric/formula keys, mixed types/errors/blanks, equal keys, ascending/descending and cross-sheet dependencies.
- A transaction that edits a dependency before sorting proves command-stage reads; stale baseRevision rejection proves revision isolation at publication.
- Relative/absolute formula movement, row identity, edit/undo/redo, zero-cache mode and reopen comparisons.
- Cancellation during key calculation and output staging, limited-memory/storage rejection, scratch cleanup and no partially published results.
- A bounded formula-heavy workload with measured key-evaluation, staging, merge and total timings. Existing scalar capacity benchmarks cannot establish formula-sort performance.

## Adversarial review

Five likely failures: using committed-head values after a preceding transaction edit; recomputing keys after rows move; losing type/error distinctions in comparison; missing the calculation cache in the operation budget; and publishing partial output after cancellation. Each needs a regression before this task is marked complete. Implementation now uses a separate stage-bound calculator, shared total comparator, immediate key-byte accounting and cache reservation within the operation budget. Regression results are recorded in `verification.md`; large-workload phase timings and exhaustive fault injection remain acceptance work.
