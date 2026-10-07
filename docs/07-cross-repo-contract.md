# Host integration contracts

> Design notes, not a claim of implemented functionality. See [implementation status](implementation-status.md) and the [current API](api.md) for supported behavior.

## Ownership and dependency boundary

A host consumes versioned `opensheet` and `@opensheetjs/*` packages through their public entry points. The OpenSheet model and commands remain authoritative; hosts must not copy implementation code or mutate internal records. Network, identity and durable storage belong to the integrating application.

This document describes future transport integration requirements. OpenSheet 0.1 does not provide a hosted persistence or collaboration service.

## Version negotiation

A future transport should explicitly negotiate coreVersion, schemaVersion, commandProtocolVersion, formulaVersion and supported capabilities. Reject incompatible writable clients; read-only loading also requires a schema the client can interpret. Unknown commands must not be silently skipped.

WorkbookFile (schema 3) is the sole public JSON document and can be passed directly to app.load/new Workbook. Binary native files carry equivalent data; old JSON and persisted engine formats are rejected. No historical migration framework is implemented. Hosts should use semantic metadata, field-selected range reads, revision-bound cursors and validated command envelopes for agent operations; see [API](api.md) and [format](storage-format.md).

## Transport responsibilities

A host may define load, submit, subscribe and getSnapshot operations. Authenticate and authorize independently of client payloads. Validate resource costs and revision preconditions. A durable acknowledgement must precede a saved-state indicator.

Retries must retain command IDs and payload identity. A durable idempotency key binds identity, workbook, command ID and canonical payload. Reusing an ID with different content must fail. Database constraints and revision compare-and-swap establish correctness; process locks alone are insufficient.

## Atomicity and recovery

Keep logical batches atomic. Large asynchronous operations should compute a complete staged result and switch it once against the expected revision. On lost acknowledgements, query or retry the original operation rather than inventing a new ID. Cross-workbook transactions are outside the current engine contract.

Client optimism should keep a confirmed snapshot and pending changes separately. On conflict, remove pending changes and reload; do not use inverse commands to overwrite another writer. Shared undo would need conditional compensation, not the instance-local undo stack.

## Compatibility checks

Test public types, schemas, errors, deterministic fixture results, import/export fidelity and revision-conflict behavior. Do not interpret current commit events as a complete replicated log. Host integrations must test disposal, plugin reinstallation and error handling without relying on private implementation APIs.
