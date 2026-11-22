# Clipboard copy batching: local diagnostic

On an Apple M4 with local headless Chromium 153, three alternating pairs copied a 1,000 × 10 numeric range using the same real Worker. The reconstructed previous one-read implementation took 28.2–35.5 ms; the current 256-cell batches took 12.6–25.0 ms. Every result matched the complete expected TSV. [Raw samples and scope](../benchmarks/clipboard-copy-batches-2026-11-22.json) are retained.

This is a small Vite development-server experiment on one machine and one data shape, not a production latency or memory guarantee. It compares repeated reads against the same workbook, uses a reconstructed baseline rather than an old build, and does not measure native clipboard operations, process peak memory or long-text-heavy ranges. Additional Worker requests can still cost more on other workloads and devices.

To reproduce, start `pnpm dev --port 5174 --strictPort` and run `node scripts/clipboard-copy-benchmark.mjs --output=benchmarks/NEW-NAME.json`. The script refuses to overwrite an existing report.
