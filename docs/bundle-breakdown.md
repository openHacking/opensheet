# Candidate bundle breakdown

These measurements describe the unreleased **0.2.0-dev.0** workspace candidate, built locally on 2026-11-22 with Node 22.19.0 and esbuild 0.25.12. They are a diagnostic for selected imports, not a size guarantee for an application.

| Synthetic entry | Output | Bytes | Gzip bytes |
| --- | --- | ---: | ---: |
| `createWorkbook`, `openWorkbook` | main JS | 543,307 | 120,318 |
| `createOpenSheet` and stylesheet | main JS | 686,212 | 160,812 |
| `createOpenSheet` and stylesheet | CSS | 9,941 | 2,663 |
| Engine Worker, bundled separately | Worker JS | 640,426 | 150,613 |
| SheetJS adapter and complete `xlsx` namespace | JS | 1,051,384 | 287,181 |

The core and editor entries do **not** include the separately bundled Worker. Esbuild does not automatically rewrite this package's `new Worker(new URL(...))` into the measured Worker bundle; a real host must configure and verify that asset. The adapter example imports the entire SheetJS namespace. A narrower host, different tree shaking, CSP shim, target, minifier or package version can produce different results. The default Zod build accounts for approximately 454 KB of uncompressed output in several entries. The table uses bytes, not KiB; gzip is calculated separately per output at level 9 and does not include HTTP headers or caching.

A separate editor entry uses dynamic `import()` for the adapter and SheetJS. Its initial JS files are 139,378 and 570,565 bytes, plus 9,941 bytes of CSS; the optional XLSX chunk is 543,834 bytes. Their separate gzip sizes are 39,962, 129,605, 2,663 and 177,202 bytes respectively. The harness walks esbuild's static output graph and checks that inputs from `xlsx` appear only outside the initial graph. The graph check does not claim that all adapter code is isolated. A separate localhost browser run described below verifies the XLSX chunk request for these outputs.

## Reproduce

From the repository root with Node 22 and pnpm 10.22.0:

```sh
pnpm install --frozen-lockfile
pnpm build
node scripts/bundle-breakdown.mjs benchmarks/NEW-bundle-report.json
# Add --csp to measure and run a separate strict-CSP build:
node scripts/bundle-breakdown.mjs benchmarks/NEW-csp-report.json --csp
# Deliberately inject a blocked operation into the engine Worker as a diagnostic control:
node scripts/bundle-breakdown.mjs benchmarks/NEW-csp-probe.json --csp-probe
```

The script refuses to overwrite the report or consumer lock. It packs six current workspace packages, installs them into a temporary consumer, records archive and entry hashes, copies the pinned SheetJS SHA-512 from the workspace lock into the consumer lock, and performs a frozen reinstall. It then minifies synthetic browser entries, records output and per-package uncompressed contribution sizes from esbuild's metafile, checks the lazy graph, and runs the resulting lazy entry in three headless browsers. It deletes the temporary consumer after the run. A working pnpm store or access to the dependency sources is required; no package is published.

The [first complete run](../benchmarks/bundle-breakdown-final-2026-11-22.json), [repeat](../benchmarks/bundle-breakdown-repeat-2026-11-22.json), and [post-rebuild run](../benchmarks/bundle-breakdown-rebuilt-2026-11-22.json) produced identical hashes for all measured output files on the same host. The first two also produced matching package archives and consumer locks; after rebuilding, the `opensheet` archive and lock differed even though output hashes matched. This does not establish cross-platform reproducibility. Consumer locks and archive hashes are retained with each report.

The [first frozen-reinstall attempt](../benchmarks/bundle-breakdown-frozen-retry-2026-11-22.json) failed because pnpm omitted integrity for the SheetJS URL. An [initial integrity-copy attempt](../benchmarks/bundle-breakdown-integrity-2026-11-22.json) produced a malformed lock because its regex captured across lines. The completed runs constrain that match and verify frozen reinstall. The workspace's SheetJS digest is a pinned checksum, not a vendor signature; see [security evidence](verification.md).

The [localhost runtime run](../benchmarks/bundle-breakdown-served-2026-11-22.json) serves the measured lazy JS, CSS and separately bundled Worker under `/embedded/sheets/`; it checks the served bytes against their measured hashes. Chromium, Firefox and WebKit each mount the editor, calculate a formula and confirm no XLSX chunk request occurs before activation. Two conversions write and read an XLSX file, load the result, reopen it and verify all three values; the XLSX chunk is requested once. The run also checks one stylesheet, the Worker request and a writable reopened workbook. These output hashes match the earlier size samples. This is localhost request observation, not a network-transfer or performance measurement.

A separate [`--csp` run](../benchmarks/bundle-breakdown-csp-verified-2026-11-22.json) builds the same synthetic entries with the documented Zod jitless shim and serves them with `script-src 'self'; worker-src 'self'; style-src 'self'`, without `unsafe-eval` or `unsafe-inline`. In Chromium, Firefox and WebKit the editor, Worker and two lazy XLSX conversions pass; external page and Worker probes each confirm that dynamic code is blocked. Document CSP events observed during the run were limited to the deliberate page probe. Worker violation events were not collected, so the absence of other Worker violations is not established. The shim changes the output hashes and sizes; this run is not a CSP test of the default bundles in the table. An [initial CSP build](../benchmarks/bundle-breakdown-csp-2026-11-22.json) failed because the isolated consumer lacked the shim's direct Zod dependency; the completed run adds the version pinned by core. A [non-CSP follow-up](../benchmarks/bundle-breakdown-nocsp-verified-2026-11-22.json) checks the original path after this change.

A subsequent [instrumented Worker run](../benchmarks/bundle-breakdown-worker-observed-2026-11-22.json) forwards `securitypolicyviolation` events from the actual engine Worker to the harness. No engine events were observed during the tested workflow. A separate [negative-control run](../benchmarks/bundle-breakdown-worker-probe-2026-11-22.json) deliberately attempts dynamic code in that Worker: Chromium and Firefox reported the blocked operation, while WebKit reported no Worker events in the one-second observation window. All three still blocked the operation and completed the workflow. The injected listener changes the Worker bytes; the negative-control build also injects the blocked operation, and neither is the uninstrumented production Worker. The separate [non-CSP run](../benchmarks/bundle-breakdown-worker-nocsp-2026-11-22.json) checks that the ordinary path still works. Absence of Worker events, especially in WebKit, does not prove absence of violations.

These measurements do not establish SSR, other bundlers, execution latency, memory, external network behavior or native devices. CSP observations are limited to this policy, the exercised workflows and observed telemetry.
