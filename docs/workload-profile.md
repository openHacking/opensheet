# Initial acceptance workload profile

DISC-01 provisional profile, 2026-10-22. No independent-user validation is claimed. Preserve the roadmap thresholds until an explicit scope ADR changes them.

## Task scripts

1. **Order worksheet:** create two sheets, enter text product IDs (including leading zeros), quantities, decimal prices and booleans; calculate line totals and SUM; edit, undo/redo, filter and sort supported scalar columns; download native JSON/binary, reopen and compare every input, formula result, style, merge and sheet identity. XLSX export must show known losses before download. Future validation and multi-key sorting join this script when implemented.
2. **Budget model:** use Budget Simulator; change D2 to 3, verify F2 = 7200, undo to D2 = 1, redo; enter cross-sheet and mixed absolute references; repeat after eviction and reopen. Formula errors and stale revisions must be actionable. Expand to named ranges and the agreed formula corpus without replacing baseline assertions.
3. **Local file review:** import a synthetic XLSX with strings, blanks, numbers, booleans, dates, formula-without-cache, merges and two sheets; inspect report; edit and export; independently open/recalculate/save in Excel and LibreOffice; reimport and compare documented supported assertions. Keep original input. A parse error, quota failure or cancelled import must leave the previous revision intact.

Existing automated anchors: `tests/e2e/playground.spec.ts`, `tests/e2e/scenes.spec.ts`, `tests/interop.test.ts`, `tests/storage.test.ts`. These are starting evidence, not completion of each future workflow.

## Deterministic dataset families

- Mixed capacity: existing Performance Lab generator and unchanged full-cell checksum/undo/dispersed-edit protocol. Preserve its seed/formula in downloaded result artifacts.
- Dense scalars: rectangular rows of text IDs, positive/negative numbers, zero, false and empty strings.
- Sparse: occupied cells in distant 64 × 32 pages, including dimension edges.
- Strings: repeated and unique long text, Unicode and unpaired surrogate preservation.
- Formulas: cross-sheet chains, fan-out, overlapping ranges, lazy branches and errors.
- Metadata: styles, dimensions, hidden axes, merges, frozen panes and history.

Each new fixture must record dimensions, seed or literal inputs, independently specified expected values and intended failure boundary. Large fixtures must be generated rather than committed as opaque binary blobs without provenance.

## Reference environment and limits

Available local diagnostic machine: Mac mini Mac16,10, Apple M4 (10 CPU cores), 16 GB memory, arm64, AC power. Browser build, OS version, command, tree revision and concurrent workload belong in each result. This machine alone does not satisfy the Windows reference-machine gate; no Windows results are available. Native screen-reader/IME trials require their own OS/browser/input-method records.

Required strict capacity gate: 1,000,000 mixed populated cells, three rounds, edit P95 ≤100 ms, complete scroll-frame P95 ≤33 ms, restore ≤10 s, complete data verification and dispersed undo/redo. 10m remains a conditional stretch target. Run capacity benchmarks sequentially with other test/build workloads stopped.

Resource limits fail explicitly and atomically. Default storage budget is 256 MiB; ordinary traversal and undo/working-set budgets are 32 MiB each; a read batch is at most 16,384 cells. XLSX conversion remains materialized. Quota/permission checks at initialization cannot promise future space or protection from browser origin eviction. Native backups must be downloaded separately.

## Evidence outside this environment

Five developer interviews, three independent integrations, Windows hardware, Excel and complete stable LibreOffice reader acceptance, real native IME and screen-reader trials, two release-capable maintainers and the RC observation period remain open. Record actual participants and consent; do not invent results or designate maintainers without availability.
