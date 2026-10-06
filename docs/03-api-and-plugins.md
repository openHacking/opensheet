# Public API and extension design

> Design notes, not a claim of implemented functionality. See [implementation status](implementation-status.md) and the [current API](api.md) for supported behavior.

## Facade and headless engine

`createOpenSheet` hosts a browser editor; `createWorkbook` from `@opensheetjs/core` creates a headless workbook. A facade owns one active workbook. Loading replaces its data and disposes old handles. Public handles must preserve validation, history and read-only semantics.

Use range methods for ordinary integrations and command schemas for generic tooling. Keep A1 addresses and zero-based exclusive numeric ranges clearly distinguished. Expose error codes rather than relying on exception strings. Snapshots are serializable data, not writable internal state.

## Plugin contracts

Plugins declare ID, version, apiVersion, dependencies and capabilities. Dependency ordering, duplicate detection, setup failure cleanup and dependent-aware removal are enforced. Current capabilities cover selection reads, workbook reads/writes, toolbar registrations and commit events. See [plugin guide](plugins.md) for exact supported APIs.

Setup should register resources. Writes during setup are not rolled back on installation failure. Resources should use cleanup hooks or returned disposal. Reloading the workbook reinstalls plugins against the new workbook. Commit handlers must avoid synchronous writes and feedback loops.

## Trust and future extensions

Capabilities prevent accidental misuse; they are not isolation. Same-page plugins can access browser privileges and must be trusted. Untrusted extensions would need a separate-origin sandbox or isolated process with validated messages, explicit data access and resource restrictions.

Custom editors, formula functions, shortcut registries, persistence providers and plugin scaffolding are future design topics. Do not invent current APIs for them. Future CLI or agent tooling can wrap the same validated command contracts rather than bypassing core mutation rules.

## Framework integration

React and Vue wrappers manage editor lifecycle and expose the same workbook APIs. initialSnapshot is mount-time input; later replacements use app.load. Components need an explicit height and disposal at unmount. Host integrations should test their framework's lifecycle behavior and remounts.

Changes to public contracts need a design note, compatibility explanation, consumer type checks and a changelog entry. Stable v1 guarantees require broader integration and backward-compatibility fixtures.
