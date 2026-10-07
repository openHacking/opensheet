# Contributing to OpenSheet

We welcome bug fixes, tests, documentation and independent plugins. Read the [API guide](docs/api.md) and [implementation status](docs/implementation-status.md) first. Numbered design documents describe long-term goals, not completed features.

## Development setup

Use Node.js 22+ and pnpm 10.22.0.

```sh
pnpm install --frozen-lockfile
pnpm dev
```

Open http://127.0.0.1:5173. The demo exposes its current public API as `window.opensheet`. Samples reset on refresh; download JSON to keep edits.

```sh
pnpm format
pnpm check
pnpm exec playwright install chromium webkit
pnpm test:e2e
OPENSHEET_BASE_PATH=/opensheet/ pnpm build
pnpm test:pages
pnpm benchmark
pnpm benchmark:browser -- --compare
pnpm preview
```

`pnpm check` checks formatting, types, unit tests, package declarations, consumer imports, schemas and the production demo build. Run browser tests for UI or file integration changes. Run `pnpm build` before the browser benchmark; it measures 10,000, 100,000 and 200,000 populated cells in headless Chromium. Use `pnpm benchmark:browser -- --write-baseline` only after reviewing a new baseline. Comparison is meaningful only on the same browser and machine. Benchmarks are diagnostics, not performance guarantees.

## Project structure

- `packages/`: package implementations; public entry points are `src/index.ts`.
- `apps/playground/`: the demo gallery, editable workbooks and browser file processing.
- `examples/`: compile-checked headless and plugin examples.
- `tests/`: unit, property, interoperability and browser tests.
- `scripts/`: builds, schemas, notices, consumer checks and benchmarks.
- `docs/`: API guides, implementation limits and architecture decisions.

The main package is published as `opensheet`; scoped packages use `@opensheetjs/*`. Version 0.1.0 is a development preview. Use workspace dependencies for local development. Do not introduce DOM, network or billing dependencies into the core.

## Making changes

Use a focused branch and include a minimal reproduction for bug fixes. New commands need validation, atomic failure, history and resource-budget coverage. Format changes need synthetic fixtures and round-trip tests. Never include private workbooks or secrets.

Propose public API, snapshot schema or command semantics changes with a design note explaining the problem, compatibility and migration. Modify source schemas, then regenerate with `pnpm build`; do not hand-edit generated schemas or `api-manifest.json`.

Explain any new dependency's necessity, size, license and maintenance. Record breaking 0.x changes in the changelog. Contributions are submitted under Apache-2.0; only contribute material you can license.

## Pull requests

Describe the concrete problem, resulting behavior, validation and remaining limits. Avoid unrelated refactors. Keep sample data synthetic or explicitly redistributable. Follow the issue and PR templates.

## Documentation and branding

Project documentation and demo UI use English. Non-English input fixtures are intentional regression coverage. Keep capability claims aligned with the implementation. Brand assets live in `docs/assets`; the demo serves copies from its public assets directory.

## Pages deployment

The production site uses `/opensheet/` as its Vite base path. `pnpm build` generates `apps/playground/dist`. GitHub Actions validates the project and deploys that artifact from `main` to GitHub Pages; pull requests only run checks. Enable the repository's Pages source as GitHub Actions. No npm publishing workflow is configured.
