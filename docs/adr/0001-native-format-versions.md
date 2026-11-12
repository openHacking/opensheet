# ADR 0001: Separate native file and engine storage versions

Status: accepted as a compatibility policy; migration implementation remains open. Date: 2026-10-30.

## Problem

A validation rule changes workbook meaning: an older reader must not silently drop it and allow invalid edits. The public JSON file, binary container and IndexedDB engine currently share some data but do not have the same version boundary. Treating them as interchangeable would make upgrades and rollback unsafe.

## Current contract

| Layer | Current identifier | Reader behavior | Compatibility scope |
| --- | --- | --- | --- |
| Native JSON WorkbookFile | `schemaVersion: 3` | Validates the schema and rejects unknown file fields | Schema 3 only; no automatic migration |
| Native binary | `OSB` signature and container byte `3`; metadata is a schema 3 WorkbookFile | Checks the container and metadata separately | Container 3 with supported metadata only |
| IndexedDB physical database | IndexedDB database version 1 and four object stores | Browser opens the existing database without a format migration | This version describes object stores, not workbook semantics |
| Persisted engine head | `engineFormatVersion: 1`, directory-based head and schema 3 encoded metadata | Accepts version 1 and the existing unversioned directory layout; rejects other versions before recovery | Current directory-based layout only; no general upgrade or downgrade path |

Schema 3 rejects unknown native fields, including nested styles, axis metadata, freeze settings and cached scalars. The explicitly open `extensions` namespace is not an enforcement boundary; an older reader can ignore its contents.

Recognizable JSON schema versions other than 3 and recognizable `OSB` container versions other than 3 return `UNSUPPORTED_VERSION`. Invalid signatures, incomplete objects and malformed current-version data remain invalid input. Detection does not prove a future format is otherwise well formed. The error directs users to a compatible reader; it does not promise that a newer workbook can be exported to schema 3 without loss.

Pre-0.1.1 persisted engine layouts and native schema versions other than 3 are **unsupported**. This policy does not delete them or infer their contents. Keep a separate backup and use a compatible older application to export data in a supported form; verify unsupported features and conversion loss before replacing the original. Browser-origin storage can be evicted and is not a backup service.

Adding the version marker does not change the directory format or workbook semantics. Existing unversioned directory heads are read without rewriting the head; the next successful atomic head publication adds version 1. A failed publication leaves the original head in place. Older applications that ignore this new field cannot be relied upon to reject future engine versions, so this marker alone does not guarantee safe downgrade.

The current reader also rejects unknown head and top-level stored-metadata fields before recovery or cleanup. This prevents unrecognized mandatory semantics or references from being silently treated as ordinary version-1 data. It does not make earlier released readers safe for downgrade, and it does not replace a future engine-format version bump.

Opening also checks the metadata envelope, schema marker and strict compressed metadata shape of each distinct revision referenced by undo or redo, before reclamation. An unsupported historical revision can otherwise be silently discarded by a later edit or history pruning. This check does not expand historical sheet axes or verify their identity relationships; full metadata decoding still happens when restoring a revision.

## Rules for future changes

1. A semantic addition such as validation rules must have an explicit file-schema version and a reader that either preserves and enforces its semantics or rejects it. Hiding mandatory semantics in generic extensions is insufficient because older readers may ignore them.
2. Container framing and engine storage must be versioned independently of the WorkbookFile schema and IndexedDB object-store version. Never infer a supported version solely from a field name.
3. Before offering an upgrade, create and verify a separate backup. Stage converted records under new identities; validate all supported features and resource limits. Publish the new head only after the complete result is durable. A crash, cancellation or quota error before publication must leave the old head readable.
4. Rollback and downgrade must be explicit. Preserve the original backup and report features that cannot be represented; do not claim lossless downgrade without a feature-indexed independent verification corpus.
5. Unsupported or partially upgraded data must remain untouched and produce an actionable error. Do not automatically create a replacement workbook or reclaim unreferenced-looking records when a head is missing.

## Not yet implemented

An internal, in-memory prototype converts fully validated schema 3 files to a provisional schema 4 with empty validation collections and rejects downgrade when rules exist. It is not a public file reader, durable engine migration or supported schema 4 format. There is no atomic multi-version upgrade protocol, compatibility fixture matrix or automatic rollback. COR-02 remains open until implementation and fault-injected upgrade, quota, interruption, rollback and independent-reader tests prove these rules.

A separate internal schema 5 experiment carries both provisional validation rules and conditional formatting; it rejects lossy downgrade to the experimental schema 4. These numbers are test artifacts, not a final version assignment or a supported file format. No engine head is published by these staging helpers.
