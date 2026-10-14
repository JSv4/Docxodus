# Portable history files (.NET)

Use `.docx` for one exact snapshot; use `.docxhistory` to carry the document and its
retained history together. Everyday storage remains host-owned and invisible to users.

```csharp
// Export from existing host-owned history. Capture happens once, even during later saves.
await using (var file = File.Create("Agreement.docxhistory"))
    await history.ExportHistoryArchiveAsync(documentId, file);

// Open independently of the original store. This object exposes read methods only.
using var input = File.OpenRead("Agreement.docxhistory");
using var document = await DocxHistoryArchive.OpenAsync(input);
var latest = await document.ExportDocxAsync();
var page = await document.ListVersionsAsync(limit: 25); // newest first; page.Next continues
var selected = await document.ExportDocxAsync(page.Versions[0].Id);
var redline = await document.CompareVersionsToDocxAsync(beforeId, afterId, diffSettings);
```

For ordinary editing, bind a stable ID to host-owned stores. A new DOCX initializes
history with one explicit checkpoint; opening an existing ID just reads its head.

```csharp
var history = new DocxVersionHistory(
    new FileHistoryBlobStore(Path.Combine(privateAppDirectory, "blobs")),
    new FileHistoryHeadStore(Path.Combine(privateAppDirectory, "heads")));
var live = history.Document(stableDocumentId); // Also accepts browser/backend adapters.
var current = await live.ReadAsync();
var saved = await live.CreateVersionAsync(requestId, current?.Head, editorDocx, metadata);
var restored = await live.RestoreVersionAsync(restoreRequestId, saved.Head, chosenVersionId, metadata);

// Resume an imported history under its original document identity and revision.
using var incoming = File.OpenRead("Agreement.docxhistory");
var imported = await history.ImportHistoryArchiveAsync(incoming);
var resumed = history.Document(imported.Archive.DocumentId);
var importedDocx = await resumed.ExportDocxAsync(imported.View.Version.Id);
```

Persist request IDs **and their exact inputs** before save/restore; reuse them unchanged
after an uncertain response. A retry returns its original result, which need not be the
latest head. Restore appends a version; load its exported bytes into the editor explicitly.
`StaleHead` means refresh and ask the user how to handle their unsaved draft.

Import validates before destination writes, then commits one absent-only exact head.
`AlreadyPresent` means the exact archive head was already present. `ImportConflict`
means a different local head exists (even a newer one): nothing is overwritten; offer
read-only opening. Custom head adapters need `IHistoryHeadInitializer`, otherwise import
fails `InitializationUnsupported` before writes. Failed imports can leave unreferenced
immutable blobs; host retention owns cleanup. Retry the same file after uncertain I/O.

`DocumentId`, `View`, and `Info.Head` identify the pinned document/version. The reader
also supports sequence/time lookup, replay, updates, and preserved operation proposals.
`CompareVersionsToDocxAsync` uses the existing DocxCompare product revision policy,
matching client `compareVersions`. The older `CompareVersionsAsync` returns a richer
raw DocxDiff comparison with caller-controlled policy; that API remains unchanged.

Await active calls before disposal. Stream input must be readable, seekable, and unchanged
until disposal; it is left open by default (`leaveOpen: false` transfers ownership).
The `byte[]` overload takes an owned copy. Reads are serialized internally at the ZIP
entry level. Output may be forward-only, must be empty, and is always left open. A failed
export may leave partial bytes: save to a private temporary file and publish it only on
success. Read-only opening/export never changes a history head or a live editor.

Use `DocxHistoryArchiveLimits` to bound file size, individual blobs/snapshots, total blobs,
metadata, graph edges, graph-blob validation reads, and conservative package-expansion work.
ZIP metadata has separate container, entry-count and manifest limits.
`PackageManifestOptions` additionally bounds each nested DOCX's ZIP/XML inspection.
Resource exhaustion and malformed content fail explicitly; no partial reader is returned.

DOCX export preserves the chosen bytes, including any existing Word comments/tracked
changes. It does **not** sanitize them, and never adds the external archive history.
An archive is neither signed nor encrypted. It preserves backend decisions without
independently certifying their acceptance/conflict policy.

See [format and integrity contract](architecture/portable_history.md) and
[real sample files](../TestFiles/HistoryArchive/README.md).
