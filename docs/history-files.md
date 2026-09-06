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
var diff = await document.CompareVersionsAsync(beforeId, afterId, diffSettings);
var redline = diff.ToRedline().DocumentByteArray;
```

`DocumentId`, `View`, and `Info.Head` identify the pinned document/version. The reader
also supports sequence/time lookup, replay, updates, and preserved operation proposals.
Comparison uses existing DocxDiff settings, including its input-revision policy.

Await active calls before disposal. Stream input must be readable, seekable, and unchanged
until disposal; it is left open by default (`leaveOpen: false` transfers ownership).
The `byte[]` overload takes an owned copy. Reads are serialized internally at the ZIP
entry level. Output may be forward-only, must be empty, and is always left open. A failed
export may leave partial bytes: save to a private temporary file and publish it only on
success. Neither operation changes a history head or a live editor.

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
