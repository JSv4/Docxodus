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

This first binding layer is package-boundary history, not fine-grained typing, automatic
concurrent-edit merging, or pending-work management. Language bindings and live log followers
are subsequent stacked layers. See [history API](history.md) for durability and retention
responsibilities and [the architecture](architecture/collaboration_and_version_history.md)
for the larger collaboration roadmap.
