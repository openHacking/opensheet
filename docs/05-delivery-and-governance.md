# Delivery and governance

> Design notes, not a claim of implemented functionality. See [implementation status](implementation-status.md) and the [current API](api.md) for supported behavior.

## Release principles

The current release is a development preview, not stable v1. Public claims must reflect tested behavior. Version 0.1.1 is published to npm; source publication and Pages deployment are separate from registry publication.

Keep the core usable independently of UI or hosted infrastructure. Public features use the same validated model and command semantics. Maintain dependency licenses and notices in distributed bundles.

## Validation gates

CI runs formatting, strict types, unit/property tests, production builds, schemas and package-consumer checks. Browser regression uses Chromium and WebKit. Generated schemas, API manifest and bundled notices must be reproducible. Add fixtures for new command semantics and format interoperability.

Performance evaluation needs environment records, realistic data, repeated trials and percentiles. Screen-reader, native IME, framework-host and real spreadsheet-reader testing remain separate acceptance tasks. Do not treat synthetic tests as manual certification.

## Change management

Bug fixes use focused pull requests. Public API, serialization and command changes need RFC/ADR notes including compatibility and migration. Record breaking 0.x changes explicitly. New dependencies require necessity, license, size and maintenance review.

Fixtures must be synthetic or clearly redistributable. Never upload sensitive user workbooks, credentials or proprietary assets. Security reporting follows the repository policy rather than public exploit disclosure.

## Milestones

Prioritize correctness, calculation isolation/incremental indexing, accessibility and integration verification before expanding extension surfaces or persistence. A stable version requires public API diffs, historical compatibility fixtures, performance baselines and real host feedback. Schedule estimates are not release guarantees.

## Repository operations

Use protected review and validation practices as maintainer access permits. Pages deploys only tested main-branch artifacts through GitHub Actions. Development workflows belong in CONTRIBUTING.md. The initial npm release is published manually; no automated registry publication or signing workflow is configured.
