# Competitive research and strategic gaps

Research date: **2026-10-22 (Asia/Shanghai)**. Local baseline: `fb76c4e`, development preview 0.1.1. This is a planning assessment, not a product certification or a competitor benchmark.

Read with [ROADMAP](../ROADMAP.md), [TODO](../TODO.md), [implementation status](implementation-status.md) and [performance evidence](performance.md).

## 1. What problem should OpenSheet win?

Developers need a spreadsheet they can embed, understand, extend and keep running without losing their users' data. End users need familiar editing and dependable calculations. Organizations need clear licensing, predictable upgrades and an honest explanation of what happens to imported files.

The proposed positioning is **a permissively licensed, local-first embeddable spreadsheet with verifiable data integrity, explicit compatibility and a stable API shared by people, applications and agents**. This is a strategic hypothesis to validate with integrations, not a claim of unique market ownership.

Three initial jobs should determine priorities:

| Job                                | Representative workflow                                                                    | Success measure                                                        |
| ---------------------------------- | ------------------------------------------------------------------------------------------ | ---------------------------------------------------------------------- |
| Embed a business worksheet         | Install in React/Vue, paste orders, validate entries, calculate totals, save and reopen    | First working integration, correct results, no lost committed edits    |
| Inspect and transform a local file | Import a customer workbook, inspect losses, edit supported data, export for another reader | Explicit feature report and independent reader checks                  |
| Automate a workbook                | Inspect a bounded range, propose changes, commit against a revision, verify output         | Bounded reads, atomic writes, clear conflicts and reproducible results |

Spreadsheet SDKs are the direct market. Data grids are adjacent alternatives for structured records. Calculation and file libraries are building blocks. Excel/Google Sheets are behavioral references, not products OpenSheet can replace comprehensively in its first stable release.

## 2. Method and limits

- Inspected local architecture, public model, formula evaluator, CI, browser configuration, adapter report implementation and checked-in benchmark artifacts.
- Reviewed official repositories and vendor documentation. Sources below were accessed on the research date; moving documentation URLs are not immutable version snapshots.
- Separate **repository evidence**, **vendor-documented capability**, **measured local observation**, and **proposed target**. None implies the others.
- No competitor package was installed or benchmarked in this research. There is no defensible speed, memory, bundle-size or fidelity ranking yet.
- Do not infer that a feature is absent because a page omits it. Edition boundaries must be rechecked against the exact package/version selected for an evaluation.
- Do not treat stars, advertised row counts or function counts as product quality. No adoption ranking or fabricated user interviews is included.
- License summaries describe official published boundaries, not a legal compatibility opinion. Any redistribution decision must inspect the actual pinned dependency license and distribution obligations.

## 3. Competitive landscape

| Product / category                                            | Verified boundary and capability                                                                                                       | What OpenSheet should learn                                        | Strategic response                                                                                          |
| ------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------- |
| **Univer — closest direct OSS SDK**                           | Apache-2.0 OSS foundation with plugin architecture, Canvas rendering and formulas; official repository separates Pro capabilities [S1] | Coherent model/commands/API and broad spreadsheet workflows        | Match essential workflows before expanding into more document types                                         |
| **Univer Pro — advanced direct benchmark**                    | Documented exchange path needs conversion backend; collaboration uses Pro client packages and backend integration [S2–S3]              | Productize import/export and recovery, not just parsing            | Offer a useful local workflow with explicit fidelity boundaries; do not claim equivalent fidelity           |
| **Handsontable — mature embedded editing benchmark**          | Grid has commercial and noncommercial/evaluation terms; HyperFormula is separately licensed [S4]                                       | Host integration and interaction consistency                       | Measure install-to-edit and keyboard task completion                                                        |
| **HyperFormula — calculation benchmark**                      | Headless engine; repository advertises 400+ functions; GPLv3/commercial dual licensing [S6–S7]                                         | Formula semantics and focused calculation API                      | Build a compatibility corpus; do not silently bundle it into the Apache distribution                        |
| **SpreadJS — commercial feature-depth benchmark**             | Vendor describes Excel-style spreadsheet, 500+ functions and advanced add-ons [S8]                                                     | Complete worksheet workflows and long-tail compatibility           | Use as a behavior/feature reference, not a promise of immediate feature parity                              |
| **Jspreadsheet CE / Pro — lightweight integration benchmark** | CE v5 repository is MIT and demonstrates worksheet configuration and typed editors; Pro has separate licensing [S9–S10]                | Low-friction embedding and useful cell editors                     | Keep first successful integration simple; verify current editions rather than copying old comparison tables |
| **Glide Data Grid — rendering benchmark**                     | MIT React Canvas grid with lazy cells and custom rendering [S11]                                                                       | Rendering latency and customization                                | Compare UI-only workloads separately from durable workbook workloads                                        |
| **AG Grid — business data-grid benchmark**                    | Community MIT; Enterprise commercial. Docs separate basic grid functionality from advanced analysis/export [S12]                       | Documentation, framework integration and structured data workflows | Avoid competing with every server-backed record-grid feature                                                |
| **Luckysheet / Luckyexcel — migration audience**              | Official repositories are archived; Luckysheet points toward Univer [S13–S14]                                                          | Existing users need an upgrade path and clear limitations          | Investigate demand before funding a migration adapter; do not use abandoned software as the sole benchmark  |
| **SheetJS CE — interoperability dependency**                  | File data tooling; official docs distinguish Pro styling/rich-text enhancements [S15]                                                  | File parsing is only one layer of fidelity                         | Keep adapter replaceable and measure our own mapping losses                                                 |

### Univer: the comparison that most changes the plan

Univer's official OSS feature list includes data validation and conditional formatting; its guides show numeric/date/list/custom-formula validation and conditional rule APIs [S16–S17]. OpenSheet needs these for everyday business worksheets, regardless of whether it ever ships collaboration.

Univer already presents agent-oriented integration in its repository [S1]. “Has an agent API” is therefore insufficient differentiation. OpenSheet must demonstrate safer bounded inspection, conflict detection and verification on actual tasks.

The current documented import/export integration requires a backend even for snapshot conversion [S2]. That creates a potential deployment advantage for OpenSheet's browser-only path, but says nothing about relative file fidelity or performance. Custom third-party conversion paths may exist; this research does not rule them out.

### Handsontable and HyperFormula: quality and licensing are separate questions

Handsontable documents keyboard/ARIA behavior and screen-reader testing [S5]. This is a useful standard for our acceptance process; Canvas plus an ARIA attribute is not equivalent to a verified accessible editing workflow.

HyperFormula is valuable as a semantic comparator. The default formula engine decision must account for distribution obligations and host licensing, alongside feature coverage and performance. Keep any optional-provider proposal behind a separate ADR and legal review of the exact integration [S6–S7]. Do not interpret GPL as simply “noncommercial only.”

### SpreadJS: use breadth to identify workflows, not to inflate scope

The vendor lists advanced spreadsheet features and optional products such as pivot tables and a collaboration server [S8]. The useful lesson is workflow completeness: select, enter, calculate, validate, inspect, export and recover. Our initial answer is a smaller supported subset with published proof. This study did not test SpreadJS fidelity, accessibility or latency.

### Jspreadsheet, Glide and AG Grid: avoid comparing unlike products

Jspreadsheet CE's current quick start includes typed input controls and worksheets [S9]. Older edition comparison pages can contradict current examples; do not import those old negative claims into our matrix.

Glide's lazy-rendering model and AG Grid's virtualized grid are relevant to the viewport. Their rendering tests do not establish durable workbook capacity, formula correctness or XLSX round trips [S11–S12]. Conversely, requiring those features in every grid comparison would unfairly penalize products aimed at a different job.

## 4. OpenSheet baseline: strengths, gaps and evidence

| Area                      | Repository evidence at baseline                                                                                         | Gap / consequence                                                                                        | Planned response           |
| ------------------------- | ----------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------- | -------------------------- |
| Atomic local storage      | Worker/IndexedDB, staged head commits, recovery, persistent history; `docs/01-architecture.md`, `tests/storage.test.ts` | Storage upgrades and browser eviction are still product risks                                            | COR-01–03, REL-01          |
| Scale                     | Strict binary-engine 1m and 10m artifacts fail at least one interaction threshold                                       | Cannot advertise either as accepted interactive capacity                                                 | PERF-01–04                 |
| Formulas                  | 13 exported built-ins in `packages/formula/src/evaluator.ts`; async aggregation in core                                 | Large gap in business formulas; synchronous and async paths need semantic agreement                      | CALC-01–07                 |
| Structural operations     | Stable axes, reference transforms, disk sorting, atomic batches                                                         | Formula-result sorting explicitly rejected; directory/axis costs remain                                  | CALC-06, PERF-03           |
| Editing                   | Canvas/DOM editor, navigation, clipboard, freeze, zoom, fill-down action                                                | Native IME and complex clipboard workflows need real acceptance                                          | UX-01–04                   |
| Accessibility             | ARIA grid/cell metadata and HTML table affordances; remaining acceptance documented                                     | No completed real assistive-technology validation                                                        | UX-02                      |
| Validation and formatting | Basic style schema and single filter specification in `packages/core/src/types.ts`                                      | Rich validation/conditional formats need model, commands, UI and file mappings                           | DATA-01–04                 |
| XLSX                      | Adapter has exact/approximated/dropped/blocked report model and strict mode                                             | Report scope is limited to detected/mapped features; independent reader acceptance absent                | IO-01–05                   |
| Native formats            | Canonical schema 3, JSON/binary streams and CRC validation                                                              | Pre-0.1.1 layouts unsupported; current-layout unversioned heads accepted; no general migration framework | COR-02                     |
| Runtime portability       | Editor requires Workers, IndexedDB and Web Locks; pure readers/formulas can run in Node                                 | “Headless” must not imply a portable full engine without storage/runtime injection                       | DX-02–04                   |
| Frameworks                | React/Vue wrappers and consumer checks                                                                                  | Host lifecycle E2E and bundler matrix outstanding                                                        | DX-02–03                   |
| Plugins                   | SDK dependencies/capabilities/cleanup and tests                                                                         | Trusted code only; not a sandbox or security boundary                                                    | EXT-01–02                  |
| Agents                    | Metadata/windows, field selection, revision-bound cursors, command envelopes                                            | No end-to-end task benchmark or shipped remote collaboration service                                     | AGENT-01–03                |
| OSS operations            | License, security/contributing docs, CI and package smoke checks                                                        | No automated npm release/signing workflow; governance needs operational evidence                         | OSS-01–05, REL-02          |
| Documentation             | Detailed current API/status plus older design notes                                                                     | `llms.txt` had stale unpublished/synchronous statements; corrected in this planning change               | DOC-01 prevents recurrence |

All TODO identifiers resolve in [TODO.md](../TODO.md). “Gap” includes unverified behavior; it is not always a claim that implementation is absent.

### Performance evidence must survive marketing pressure

The strict [1m artifact](../benchmarks/capacity-binary-1m.json) records edit P95 **65.2 ms**, scroll P95 **35.3 ms**, restore approximately **133 ms**. The strict [10m artifact](../benchmarks/capacity-binary-10m.json) records **197 ms / 52.8 ms / 217 ms** respectively. Each stopped after its first failed round; the protocol requires three passing rounds with edit ≤100 ms and scroll ≤33 ms. These are checked-in October 7 observations, not new runs performed for this research.

The earlier 6.87m acceptance belongs to a previous engine. Hot-cell diagnostics, compressed file size and main-document heap cannot replace current dispersed-edit acceptance or total-process memory measurements. Keep failed results in the public evidence set.

## 5. Comparison protocol to implement

Use pinned releases of OpenSheet, Univer OSS, Jspreadsheet CE and Glide first. Add Handsontable/SpreadJS only with a usable evaluation/license and record edition restrictions. Include HyperFormula in a separate calculation suite. Do not distribute licensed competitor binaries with the benchmark harness.

| Track                | Shared workload                                                                              | Measurement                                                       | Fairness rule                                                          |
| -------------------- | -------------------------------------------------------------------------------------------- | ----------------------------------------------------------------- | ---------------------------------------------------------------------- |
| Rendering/editing    | 10k/100k/1m populated cells, same typed values, viewport and navigation path                 | Ready-to-edit, edit-to-paint P50/P95/P99, cold scroll, long tasks | Report UI-only and durable acknowledgements separately                 |
| Calculation          | Small business sheets, fan-out graph, deep chain, overlapping ranges, cross-sheet references | Correctness, full/incremental time, cancellation and memory       | Unsupported syntax is a capability result, never a fast successful run |
| Interoperability     | Redistributable synthetic workbooks with feature manifest                                    | Semantic diffs, visual review, known loss detection               | Same original files and independent reader versions                    |
| Developer experience | Vanilla, React, Vue, client-only SSR route, strict CSP host                                  | Setup steps, errors, bundle constituents, successful disposal     | Use supported public APIs and documented production build              |
| Accessibility        | Select/edit/copy/error correction/sheet switching without a mouse                            | Task completion with NVDA and VoiceOver                           | Record exact OS, browser, reader and tester; no inferred certification |

Pin hardware, OS, power mode, browser, viewport, package versions, commit, seed and feature settings. Run products sequentially on an idle machine; distinguish cold and warm runs. Publish raw samples, failures, exclusions and harness source. No weighted “winner” score until target users select the weights. The benchmark implementation is future work, tracked by PERF-04.

## 6. Strategic decisions and opportunity tests

| Proposed decision                                                   | Why                                                         | Evidence that could change it                                                           |
| ------------------------------------------------------------------- | ----------------------------------------------------------- | --------------------------------------------------------------------------------------- |
| Keep default core and essential workflow Apache-2.0                 | Predictable embedding is central to adoption                | Maintainer decision through public governance, not an accidental dependency change      |
| Prioritize trust and ordinary spreadsheet work before collaboration | Existing data/performance/UX gaps directly block adoption   | Multiple committed design partners require collaboration as the primary job             |
| Keep local file operations and report every detected loss           | Fits offline/privacy-sensitive embedding                    | Representative files require capabilities impossible within chosen browser budget       |
| Improve own formula semantics first                                 | Current engine is small and permissively distributed        | A measured provider comparison justifies cost, obligations and architectural complexity |
| Use agents through public commands                                  | Avoid divergent behavior and unverifiable mutations         | No exception: new capabilities should extend shared contracts                           |
| Delay multi-document office suite ambitions                         | A small team cannot maintain several mature editors at once | Sustained maintainers, funding and validated demand for another document type           |

Before treating positioning as validated, obtain five developer interviews across at least three independent organizations/projects and three hands-on integrations. Record current tool, failure point, required functions/files, offline needs, upgrade expectations and willingness to contribute. These are proposed discovery targets, not interviews already conducted. Prioritize observed task blockers over general enthusiasm.

Potential opportunities: trusted offline worksheets, transparent compatibility reports, audit-friendly automation, and a migration path for legacy embedded sheets. None is proven exclusive. Check alternatives with the same task script before making comparative claims.

## 7. Source register

All sources accessed 2026-10-22. Official documentation is evidence of documented product behavior; hands-on verification remains outstanding.

| ID  | Primary source                                                                                        | Used for / boundary                                                         |
| --- | ----------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------- |
| S1  | [Univer repository](https://github.com/dream-num/univer)                                              | Apache foundation, architecture, OSS/Pro split, agent positioning           |
| S2  | [Univer import/export](https://docs.univer.ai/guides/sheets/features/import-export)                   | Current documented backend exchange path; page displayed v1.2.0             |
| S3  | [Univer collaboration](https://docs.univer.ai/guides/sheets/features/collaboration)                   | Pro packages and backend integration; moving page version                   |
| S4  | [Handsontable FAQ](https://handsontable.com/knowledge-base/)                                          | Grid licensing and separate formula product                                 |
| S5  | [Handsontable accessibility](https://handsontable.com/docs/javascript-data-grid/accessibility/)       | Accessibility process; page displayed 18.1.2                                |
| S6  | [HyperFormula repository](https://github.com/handsontable/hyperformula)                               | Headless engine and advertised function breadth                             |
| S7  | [HyperFormula license text](https://github.com/handsontable/hyperformula/blob/master/LICENSE.txt)     | GPLv3/commercial distribution boundary                                      |
| S8  | [SpreadJS product page](https://developer.mescius.com/spreadjs)                                       | Commercial feature breadth, functions and add-ons; vendor claims            |
| S9  | [Jspreadsheet CE repository](https://github.com/jspreadsheet/ce)                                      | MIT, CE v5 examples and typed editors                                       |
| S10 | [Jspreadsheet license management](https://jspreadsheet.com/docs/license)                              | Pro/extensions licensing; no CE inference from historical tables            |
| S11 | [Glide Data Grid repository](https://github.com/glideapps/glide-data-grid)                            | MIT, Canvas, lazy rendering and React interface                             |
| S12 | [AG Grid editions](https://www.ag-grid.com/javascript-data-grid/community-vs-enterprise/)             | Community/Enterprise feature and license distinction; page displayed 36.2.0 |
| S13 | [Luckysheet repository](https://github.com/dream-num/Luckysheet)                                      | Archived October 30, 2025; successor direction                              |
| S14 | [Luckyexcel repository](https://github.com/dream-num/Luckyexcel)                                      | Archived July 8, 2025                                                       |
| S15 | [SheetJS CE overview](https://docs.sheetjs.com/docs/)                                                 | File tooling and CE/Pro boundary                                            |
| S16 | [Univer data validation](https://docs.univer.ai/guides/sheets/features/data-validation)               | Rule types and public API                                                   |
| S17 | [Univer conditional formatting](https://docs.univer.ai/guides/sheets/features/conditional-formatting) | Conditional rule model and public API                                       |

## 8. Adversarial review of this plan

| Likely failure                                                      | Correction incorporated                                                             | Remaining verification                                      |
| ------------------------------------------------------------------- | ----------------------------------------------------------------------------------- | ----------------------------------------------------------- |
| Market comparison confuses free source with permissive distribution | Editions/licenses separated; HyperFormula not assumed Apache-compatible by default  | Inspect pinned licenses when evaluating providers           |
| Large cell count hides failed real interaction                      | Both failing binary-engine artifacts retained; no capacity victory claimed          | Rerun on release commit and declared machines               |
| Roadmap becomes an unlimited Excel clone                            | Narrow target jobs, gate-based releases, post-v1 advanced scope                     | Interviews may change which workflows are essential         |
| A plan declares a feature missing or complete without evidence      | Current source/limits distinguish implemented, partial and unverified               | Actual competitor trials and integration tests still needed |
| Calendar and checkboxes create false certainty                      | Effort bands, owners-to-assign, explicit dependencies and evidence-based completion | Maintainer capacity and pilot commitments remain unknown    |
