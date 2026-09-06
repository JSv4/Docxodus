# History clients and log-driven rendering

The core history service remains the sole replay/publication implementation. Client bindings
use `Docxodus.Internal.HistoryClientOps` with host-supplied `IHistoryBlobStore` and
`IHistoryHeadStore` adapters. No networking, server, subscription, timer, or transit layer is
provided. Hosts decide when to deliver a new head or ask a client to refresh.

The version-1 client request includes `schemaVersion: 1`, `operation`, and `documentId`.
Operations are `read`, `updates`, `create`, `list`, `get`, `export`, `materialize`, `replay`,
`resolveTime`, and `restore`. Additional fields are `expectedHead`, `versionId` (also
the list cursor), `metadata`, `sequence`, `cutoff`, `limit`, and `maxEntriesToScan`.
Create receives DOCX bytes separately from its JSON request. Export/materialize/replay
return base64 bytes in their JSON result, which the language wrapper decodes.

Publication revision, content sequence, and epoch are **decimal strings** at this client
boundary, so JavaScript never rounds a 64-bit position. Blob lengths, schema versions,
and limits are numbers. The existing durable record codecs are unchanged. Requests are
bounded to 512 Ki UTF-16 characters and reject missing required, unknown, duplicate,
null-required, and malformed fields. Generated JSON metadata supports trimmed WASM.

Responses contain `success` and the relevant `view`, `update`, `version`, `page`, `sequence`, or
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

Methods are `read`, `readChangesSince`, `createVersion`, `listVersions`, `getVersion`, `exportVersion`, `materialize`,
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

The Python client exposes `read`, `read_changes_since`, `create_version`, `list_versions`, `get_version`,
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
concurrent-edit merging, or pending-work management. See [history API](history.md) for durability and retention
responsibilities and [the architecture](architecture/collaboration_and_version_history.md)
for the larger collaboration roadmap.

## MCP

Launch the existing server with `DOCXODUS_HISTORY_ROOT=/host-owned/matter-history` to enable
`docxodus_history`. It reuses existing MCP dispatch, with no new transport or subscription.
The directory contains the same protected `blobs`/`heads` layout as Python's file binding.
Host configuration controls storage and retention; tool arguments cannot widen it.

Open a scoped document with `docxodus_open`, then call `docxodus_history` with its `sessionId`
and an `action` from the shared operation list above. Do not send `schemaVersion`, `operation`,
or `documentId`: the binding assigns these, using the session's canonical scoped location as
the history identity. Other fields and results match the version-1 facade. Positions are
decimal strings; `versionId` and `expectedHead` are the exact objects returned by prior calls.
Saving a copy to another location does not change the session's original history identity;
open the copy to use that location's separate history.

`create` captures clean current session bytes without saving the source file. Pass `metadata`
and the exact `expectedHead` (null only for first publication). `restore` requires the target
`versionId`, metadata, and exact expected head, and appends history only: it leaves both the
open editing session and source file untouched. Hosts authenticate the supplied author.

The extra `render` action takes exactly one `sequence` or `cutoff`; it reconstructs the
checkpoint through the shared service and converts it to HTML with the existing renderer,
returning `{ success: true, sequence, html }`. For example:

```json
{ "sessionId": "<open-session-capability>", "action": "render", "cutoff": "2026-01-01T12:30:00Z" }
```

Use `updates` with the last accepted `expectedHead` to retrieve the ordered log described
below. An unknown/closed session or foreign-document version cannot read another history.
There is no implicit follow loop, source-file write, or live-session reset.

## Ordered live log updates

`ReadChangesSinceAsync` (.NET), `readChangesSince` (npm), and `read_changes_since` (Python)
accept the exact last accepted `HistoryHead` and return one captured `DocxHistoryUpdate`.
Each update contains that `after` head, its new `view`, an ascending contiguous list of
`entries` (immutable commit reference, commit record, and recorded author/time metadata),
and `reset`. Sequence determines order even when timestamps move backwards. Metadata-only
labels advance the publication head without creating content entries. A repeated accepted
head returns an empty tail. First join passes `null`/`None` and starts at the latest checkpoint
with an empty historical tail and `reset: true`.

Both immutable version ancestry and content-commit ancestry must extend the accepted head.
Rewinds, same-revision forks, same-content version branches, broken commit links, and absent
history fail explicitly. The scan budget covers traversed version ancestors plus content
commits; each can require several bounded metadata reads. The update call reads no snapshot
or effect bytes. Hosts must retain all metadata needed to establish ancestry.

Hosts invoke refresh when their own notification/log delivery mechanism says there is new
work. No library timer, subscribe loop, socket, transport, or service is created. A reset means
the epoch changed (for example a restore); a client must preserve local pending work separately
before installing that checkpoint. This API does not authorize dropping local edits, rebase
stale intents, or implement last-writer-wins. Exact-head publication failures keep the caller's
candidate untouched. Use the returned view's exact version reference to load its snapshot for
rendering; that reference remains stable even if newer versions publish during the load.

## Live browser viewer

```ts
import { initialize, openDocxHistory, createHistoryViewer } from 'docxodus/embed';
await initialize();
const history = openDocxHistory(hostStorage);
const viewer = await createHistoryViewer('#document', history, 'contract');

// Call from YOUR existing log/notification delivery mechanism. Nothing subscribes automatically.
const update = await viewer.refresh();
console.log(update.entries.map(e => [e.commit.sequence, e.metadata.createdAt]));

await viewer.showTime('2026-01-01T12:30:00Z'); // or showSequence('42')
await viewer.refresh(); // catches up its verified live head, but keeps the historical view visible
await viewer.resume();  // explicitly return to the latest verified live state
const bytes = viewer.exportDisplayed(); // independent copy of the displayed checkpoint
viewer.destroy();
history.close(); // await outstanding client operations first
```

`createHistoryViewer` is an opt-in, read-only collaboration surface. It renders complete
accepted DOCX checkpoints through the existing scoped viewer pipeline. It does not replace
an editing session or discard local edits. Publish edited candidate bytes with an exact
expected head through the normal client API; a stale candidate requires host/caller
reconciliation. No conflict-transform engine, keystroke recorder, network, timer, or
subscription is introduced here.

`head` reports the last verified live publication; `sequence` reports the displayed content
position, which can be older during a historical preview. `following` distinguishes those
modes. Refreshes and preview/resume calls are serialized in invocation order. Duplicate
notifications do not reread snapshots or rerender. A content-equivalent tail still advances
the displayed sequence and exact downloadable bytes. Restore epochs install a new frame.
Historical previews remain visible during log catch-up until `resume()`.

Snapshot and rendering failures leave the previously accepted head and visible frame intact.
A later refresh can retry from that head. Destruction prevents pending work from painting into
the removed mount; it does not delete or close shared history storage. The viewer retains the
live checkpoint and, while previewing, the displayed historical checkpoint. Full-checkpoint
loading/rendering and package-boundary publication can scale with document size; this is not
the architecture's future incremental typing performance path.
