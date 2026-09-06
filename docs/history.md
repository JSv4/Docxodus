# Named version history (.NET)

`Docxodus.History.DocxVersionHistory` captures exact DOCX versions over host-owned storage. It currently supports create, read, paginated list, get, export, lazy semantic/native-redline comparison, and non-destructive restore. Content-changing versions produce the same reversible package contributions used by the package fallback. Live session recording, sequence/time reconstruction, collaboration, and browser/agent bindings follow in the [implementation plan](architecture/collaboration_and_version_history.md).

```csharp
using Docxodus.History;

var blobs = new FileHistoryBlobStore("history/blobs");
var heads = new FileHistoryHeadStore("history/heads");
var history = new DocxVersionHistory(blobs, heads);

var current = await history.ReadAsync("document-123", cancellationToken);
var saved = await history.CreateVersionAsync(
    "document-123", current?.Head, suppliedDocxBytes,
    new DocxVersionMetadata
    {
        Author = authenticatedUserId,
        CreatedAt = DateTimeOffset.UtcNow,
        Label = "Negotiation draft",
        Message = "Updated closing conditions",
        ApplicationMetadata = new Dictionary<string, string> { ["matter"] = matterId },
    }, cancellationToken);

byte[] exactDownload = await history.ExportVersionAsync(
    "document-123", saved.Version.Id, cancellationToken);
var page = await history.ListVersionsAsync("document-123", limit: 25,
    cancellationToken: cancellationToken);
```

Use `MemoryHistoryBlobStore` and `MemoryHistoryHeadStore` for process-local use or tests. Custom adapters implement `IHistoryBlobStore` and `IHistoryHeadStore`; their contracts also permit a database/object-store host. Persist the entire immutable blob graph before the atomic head compare-and-swap. Do not implement the CAS as an unprotected read followed by a write.

## Publication and identity

`CreateVersionAsync` requires the exact previously read `HistoryHead`; null initializes an absent document. A stale expectation raises `DocxHistoryException` with `StaleHead`. Do not silently retry with a new expectation: a host should decide whether the submitted document still represents the desired next version. Every successful create produces a distinct version ID. After an uncertain storage response, inspect the head/list before deciding whether to issue a new create; operation-level retry deduplication is not yet implemented.

A version ID is its immutable manifest's `HistoryBlobReference` (SHA-256 plus byte length), not a mutable filename or array index. It retains a parent link, exact snapshot, creation nonce, host metadata, and content-log sequence. Raw snapshot SHA-256 binds the downloadable bytes; the independent ordered OPC digest identifies package content. No provenance is inserted into the DOCX.

Initial capture establishes sequence 0. Naming unchanged content increments only `HistoryHead.Revision`, even if ZIP timestamps/compression differ. Changed content appends one import commit, its reversible package contribution, and its named version; one head CAS publishes them together. A failed blob write or CAS leaves the prior history visible. Unreferenced blobs may remain and belong to the host's retention/cleanup policy.

Inputs are captured before awaiting host storage. The service does not mutate a `DocxSession` or the supplied byte array. These full-package captures and inspections belong at version/import boundaries, not on the typing path.

## Reading and comparison

`ReadAsync` validates the current document and the state/version/commit tip relationships. `ListVersionsAsync` starts at that published head and returns newest-first pages (1–100 records); pass a non-null `Next` to continue the same immutable chain while newer versions publish. A null `Next` means the page reached the end.

`GetVersionAsync`, `ExportVersionAsync`, and explicit list cursors also accept retained branch references belonging to the same document. A reference is neither authorization nor proof that a version is on the current published branch. The host must authorize access before invoking these APIs. Cross-document references are rejected.

`ExportVersionAsync` verifies exact length/SHA-256 and bounded DOCX/content identity before returning the original stored bytes. `CompareVersionsAsync` verifies both snapshots and returns the existing lazy `DocxDiffComparison`; use `GetSemanticChanges()` or `ToRedline()` as needed. It retains DocxDiff's compatibility warnings, pre-existing-revision policy, and unsupported-feature limitations. A redline is a review artifact, not a lossless package patch.

## Restore

```csharp
var restored = await history.RestoreVersionAsync(
    "document-123", previewedHead, selectedVersionId,
    new DocxVersionMetadata
    {
        Author = authenticatedUserId,
        CreatedAt = DateTimeOffset.UtcNow,
        Message = "Restored the approved draft",
    }, cancellationToken);
```

For `V1 → V2 → V3`, restoring V1 creates V4 with parent V3 and `RestoredFrom = V1`; V1–V3 remain untouched and exportable. The service verifies the target's exact bytes before writing new metadata, then atomically publishes the restore commit, new version, and state. It reuses V1's exact snapshot reference without reserializing it. Content sequence and epoch each advance, including when explicitly restoring already-current content. Missing/corrupt/foreign targets and stale previewed heads fail without publishing a reset; any metadata written before a failed CAS remains unreferenced.

Restore changes durable history, not an existing live `DocxSession`. Hosts must notify open clients of the new epoch, install the restored state through their lifecycle, invalidate derived caches, and preserve pre-reset pending edits as an explicit recoverable branch. The core service does not silently replay those edits or discard them; the integrated editor/client lifecycle remains a later layer.

The filesystem adapters require a protected local directory with working exclusive file sharing and atomic rename semantics. They do not coordinate distributed/network filesystems. Never delete active `.lock` files. The host owns authorization, aggregate quotas, retention, and filesystem-dependent power-loss durability; readers verify content and report corruption or missing history explicitly.
