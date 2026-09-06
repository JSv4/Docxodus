# History clients and log-driven rendering

The core history service remains the sole replay/publication implementation. Client bindings
use `Docxodus.Internal.HistoryClientOps` with host-supplied `IHistoryBlobStore` and
`IHistoryHeadStore` adapters. No networking, server, subscription, timer, or transit layer is
provided. Hosts decide when to deliver a new head or ask a client to refresh.

The version-1 client request includes `schemaVersion: 1`, `operation`, and `documentId`.
Operations are `read`, `create`, `list`, `get`, `export`, `materialize`, `replay`,
`resolveTime`, and `restore`. Additional fields are `expectedHead`, `versionId` (also
the list cursor), `metadata`, `sequence`, `cutoff`, `limit`, and `maxEntriesToScan`.
Create receives DOCX bytes separately from its JSON request. Export/materialize/replay
return base64 bytes in their JSON result, which the language wrapper decodes.

Publication revision, content sequence, and epoch are **decimal strings** at this client
boundary, so JavaScript never rounds a 64-bit position. Blob lengths, schema versions,
and limits are numbers. The existing durable record codecs are unchanged. Requests are
bounded to 512 Ki UTF-16 characters and reject missing required, unknown, duplicate,
null-required, and malformed fields. Generated JSON metadata supports trimmed WASM.

Responses contain `success` and the relevant `view`, `version`, `page`, `sequence`, or
`bytes` field. Expected domain/argument failures return `success: false`, `errorCode`, and
`message`; unexpected host-storage failures propagate. Cancellation before publication
returns `Canceled`; a successful atomic head publication is not later reported canceled.

Timestamps are recorded metadata, not ordering authority: sequence orders accepted content
changes. `resolveTime` finds a content sequence; `materialize` or `replay` returns its DOCX
for the existing rendering/session APIs. Opening a historical view never changes a shared
head. Comparison composes two exact exports with the existing DocxDiff client API.

## Browser and npm

```ts
import { initialize, openDocxHistory, createMemoryHistoryStorage, openDocxSession } from 'docxodus';
await initialize();
const storage = createMemoryHistoryStorage(); // or implement HistoryStorage for your host
const history = openDocxHistory(storage);
const saved = await history.createVersion('contract', null, docxBytes, {
  author: 'application-user-id', createdAt: '2026-01-01T12:00:00Z', label: 'Draft',
});
const sequence = await history.resolveSequenceAtTime('contract', '2026-01-02T00:00:00Z');
const session = openDocxSession(await history.materialize('contract', sequence));
console.log(session.project().markdown);
session.close();
history.close(); // closes the binding, not the host's retained data
```

`HistoryStorage` supplies asynchronous `readBlob`, `putBlob`, `readHead`, and atomic
`advanceHead(documentId, expectedHead, stateReference)` callbacks. Missing blobs/heads and
failed compare-and-swap return `null`. Successful CAS returns its incremented revision plus
the published state reference, just like `IHistoryHeadStore`. Storage callbacks must settle;
this binding does not impose a timeout or network-cancellation policy. Host exceptions propagate.
`DocxHistoryError.code` preserves domain error codes such as `StaleHead` and `PayloadMismatch`.

The memory reference adapter verifies SHA-256 and copies inputs/outputs. Multiple clients may
share it, but a page/process restart loses it. Supply durable storage for restart recovery.
`close()` rejects while calls are active; await outstanding calls, then close. A custom WASM
loader calls `installHistoryStorageImports(runtime.setModuleImports)` before opening a
`DocxHistoryClient` over `exports.DocxodusWasm.HistoryBridge`. Normal npm `initialize()` does
this automatically. The existing worker RPC does not yet expose these host callbacks.

Methods are `read`, `createVersion`, `listVersions`, `getVersion`, `exportVersion`, `materialize`,
`replay`, `resolveSequenceAtTime`, and `restoreVersion`. For comparisons, export both versions
and call the existing `docxDiffCompareProducts` API. The package-boundary WASM bridge uses
base64 for asynchronously read/exported bytes, incurring temporary allocation overhead;
it is not a low-latency keystroke path.

## Python

```python
from docx_scalpel import open_history, DocxVersionMetadata, convert_docx_to_html

with open_history('/host-owned/matter-history') as history:
    saved = history.create_version('contract', None, docx_bytes,
        DocxVersionMetadata('application-user-id', '2026-01-01T12:00:00Z', label='Draft'))
    sequence = history.resolve_sequence_at_time('contract', '2026-01-02T00:00:00Z')
    html = convert_docx_to_html(history.materialize('contract', sequence))
```

The Python client exposes `read`, `create_version`, `list_versions`, `get_version`,
`export_version`, `materialize`, `replay`, `resolve_sequence_at_time`, and `restore_version`.
Frozen value types use snake_case attributes and arbitrary-precision Python integers;
Int64 bounds are enforced before positions are encoded as decimal strings. Timestamp strings
retain their exact recorded precision. `DocxHistoryError.code` carries core domain errors;
host/process failures remain `DocxodusTransportError`. Compare exact exports with the existing
`docx_diff_compare_products` API, or open them with `open_session` for headless editing.

The existing local stdio host owns the C# adapters. An explicit root creates/uses its `blobs`
and `heads` directories, subject to the [filesystem adapter contract](history.md); multiple
clients/processes may cooperate over that protected local directory. With no root,
`open_history()` creates a private ephemeral memory history. Use the context manager or
call `close()`; closing never removes persisted data. After a host restart, reopen the root.
An old client fails against its original dead process rather than attaching its numeric handle
to an unrelated newly opened history. No new transport/server has been added.

This binding layer is package-boundary history, not fine-grained typing, automatic
concurrent-edit merging, or pending-work management. Live log followers
are subsequent stacked layers. See [history API](history.md) for durability and retention
responsibilities and [the architecture](architecture/collaboration_and_version_history.md)
for the larger collaboration roadmap.
