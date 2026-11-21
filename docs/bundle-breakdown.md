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

A separate editor entry uses dynamic `import()` for the adapter and SheetJS. Its initial JS files are 139,378 and 570,565 bytes, plus 9,941 bytes of CSS; the optional XLSX chunk is 543,834 bytes. Their separate gzip sizes are 39,962, 129,605, 2,663 and 177,202 bytes respectively. The harness walks esbuild's static output graph and checks that inputs from `xlsx` appear only outside the initial graph. This establishes a property of the synthetic build graph, not a browser-network observation or a claim that all adapter code is isolated.

## Reproduce

From the repository root with Node 22 and pnpm 10.22.0:

```sh
pnpm install --frozen-lockfile
pnpm build
node scripts/bundle-breakdown.mjs benchmarks/NEW-bundle-report.json
```

The script refuses to overwrite the report or consumer lock. It packs six current workspace packages, installs them into a temporary consumer, records archive and entry hashes, copies the pinned SheetJS SHA-512 from the workspace lock into the consumer lock, and performs a frozen reinstall. It then minifies synthetic browser entries, records output and per-package uncompressed contribution sizes from esbuild's metafile, and checks the lazy graph. It deletes the temporary consumer after the run. A working pnpm store or access to the dependency sources is required; no package is published.

The [first complete run](../benchmarks/bundle-breakdown-final-2026-11-22.json), [repeat](../benchmarks/bundle-breakdown-repeat-2026-11-22.json), and [post-rebuild run](../benchmarks/bundle-breakdown-rebuilt-2026-11-22.json) produced identical hashes for all measured output files on the same host. The first two also produced matching package archives and consumer locks; after rebuilding, the `opensheet` archive and lock differed even though output hashes matched. This does not establish cross-platform reproducibility. Consumer locks and archive hashes are retained with each report.

The [first frozen-reinstall attempt](../benchmarks/bundle-breakdown-frozen-retry-2026-11-22.json) failed because pnpm omitted integrity for the SheetJS URL. An [initial integrity-copy attempt](../benchmarks/bundle-breakdown-integrity-2026-11-22.json) produced a malformed lock because its regex captured across lines. The completed runs constrain that match and verify frozen reinstall. The workspace's SheetJS digest is a pinned checksum, not a vendor signature; see [security evidence](verification.md).

These measurements do not test runtime loading, CSP, SSR, other bundlers, execution latency, memory or native devices. Existing [esbuild consumer evidence](../benchmarks/esbuild-editor-keyboard-csp-retry-2026-11-12.json) tests a different strict-CSP build and should not be treated as execution of these measured files.
