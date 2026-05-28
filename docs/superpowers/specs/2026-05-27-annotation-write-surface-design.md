# Annotation Write Surface for `DocxSession`

**Date:** 2026-05-27
**Branch:** `feat/annotation-write-surface`
**Status:** Design — awaiting implementation plan.

## Problem

`DocxSession` exposes a rich read surface for Docxodus annotations
(`FindByAnnotation`, `FindByLabel`, `FindByBookmark`, `ListAnnotations`), but
no write surface. Agents and clients running over the WASM bridge or the
Python `docx-scalpel` package can discover annotations but cannot create,
update, or delete them — the only workaround is the raw-XML escape hatch
(`Raw.InsertXml`), which bypasses `AnnotationManager`'s bookmark allocation,
custom-XML serialization, and id-collision checks.

The .NET core does have full annotation CRUD via
`Docxodus.AnnotationManager`, but that API takes a `WmlDocument` (immutable
byte-array wrapper) and clones the document per call. That model is
incompatible with `DocxSession`'s in-place editing of a live open package and
its undo ring — every other tier (text CRUD, structural, formatting, tables,
raw) mutates `_doc!` directly and snapshots the package into the undo ring.

This spec adds a write surface that:
1. Lives on `DocxSession` and mutates the live `WordprocessingDocument`.
2. Routes through the existing `Docxodus.Internal.DocxSessionOps` facade so
   that the WASM bridge and the Python stdio host both pick up the new ops
   from one place.
3. Threads through to typed methods in the npm `@docxodus/wasm` package and
   the `docx-scalpel` Python package.

## Goals

- **Anchor-addressed.** Every write op targets an existing markdown-projection
  anchor (`{#kind:scope:unid}`) plus an optional character span within that
  anchor's block. No new targeting modes; callers reuse `Grep` / `FindByText`
  / `FindByKind` to obtain anchors.
- **Four ops:** `AddAnnotation`, `RemoveAnnotation`, `UpdateAnnotation`,
  `MoveAnnotation`.
- **One return shape.** `EditResult` envelope, identical to every other
  `DocxSession` mutation, with one new optional field (`AnnotationId`).
- **Undo/redo for free.** Each mutation is wrapped by the existing
  `_history.RecordPreOp(TakeSnapshot())` pattern; no annotation-specific undo
  logic.
- **No `AnnotationManager` churn.** The existing public `WmlDocument`-shaped
  `AnnotationManager` API is unchanged; the live-session path is a separate
  internal helper.

## Non-goals

- No alternative targeting modes (search text, paragraph indices, table
  coordinates, existing-bookmark name). Reachable via anchor discovery.
- No bulk import/export of annotations through the session API. The existing
  `ExternalAnnotationManager` covers that use case at the `WmlDocument`
  layer.
- No page-info recomputation on write. New annotations land with
  `PageInfoStale = true`, same as the existing manager's behavior.
- No changes to the .NET CLI tools (`tools/redline/`, `tools/docx2html/`,
  `tools/docx2oc/`) or the web demo (`web/DocxodusWeb/`).

## Architecture

```
┌─ DocxSession.cs ─────────────────────────────────────────────────┐
│   public EditResult AddAnnotation(anchorId, span, annotation)    │
│   public EditResult RemoveAnnotation(annotationId)               │
│   public EditResult UpdateAnnotation(annotationId, update)       │
│   public EditResult MoveAnnotation(annotationId, anchorId, span) │
│        │                                                         │
│        │ resolves anchor → AnchorTarget                          │
│        │ records pre-op snapshot for undo                        │
│        ▼                                                         │
└────────│─────────────────────────────────────────────────────────┘
         │
┌─ Internal/AnnotationOps.cs ──────────────────────────────────────┐
│   static EditResult Add(doc, anchor, span, annotation)           │
│   static EditResult Remove(doc, annotationId)                    │
│   static EditResult Update(doc, annotationId, update)            │
│   static EditResult Move(doc, annotationId, anchor, span)        │
│        │                                                         │
│        │ resolves AnchorTarget → XElement + part                 │
│        │ uses RunTextMap to locate span boundaries               │
│        │ inserts/removes <w:bookmarkStart>/<w:bookmarkEnd>       │
│        ▼                                                         │
└────────│─────────────────────────────────────────────────────────┘
         │
┌─ Internal/AnnotationsCustomXml.cs ───────────────────────────────┐
│   GetOrCreatePart / ReadAll / Write / Remove / FindById          │
│   ← extracted from AnnotationManager.cs (private helpers today)  │
│   ← also reused by the existing AnnotationManager public API     │
└──────────────────────────────────────────────────────────────────┘
```

The facade `Docxodus.Internal.DocxSessionOps` gains four new static methods
(one per op). The WASM `[JSExport]` shells in `DocxSessionBridge.cs` and the
Python NDJSON dispatcher in `tools/python-host/Dispatcher.cs` add one
one-liner each pointing at those facade methods.

## Public API (core, .NET)

```csharp
namespace Docxodus;

// New partial-update record. null fields = leave unchanged. Bookmark name
// and id are NOT updatable — they are identity. MetadataPatch is a per-key
// merge: keys with null values are removed from the existing dict.
public sealed record AnnotationUpdate
{
    public string? LabelId { get; init; }
    public string? Label { get; init; }
    public string? Color { get; init; }
    public string? Author { get; init; }
    public IReadOnlyDictionary<string, string?>? MetadataPatch { get; init; }
}

// New error codes appended to EditErrorCode.
public enum EditErrorCode
{
    // ... existing codes ...
    DuplicateAnnotationId,
    AnnotationNotFound,
    EmptyAnnotationSpan,
}

// New optional field on the existing EditResult.
public sealed class EditResult
{
    // ... existing fields ...
    public string? AnnotationId { get; init; }
}

public sealed class DocxSession
{
    /// <summary>
    /// Annotate a range inside the block addressed by <paramref name="anchorId"/>.
    /// When <paramref name="span"/> is null, the bookmark wraps every inline run
    /// of the block. When <paramref name="annotation"/>.Id is null/empty a
    /// 16-char hex id is generated. BookmarkName, Created, AnnotatedText, and
    /// PageInfoStale on the passed-in annotation are always set by this method;
    /// any caller-supplied values for those fields are overwritten.
    /// </summary>
    public EditResult AddAnnotation(
        string anchorId,
        CharSpan? span,
        DocumentAnnotation annotation);

    /// <summary>Delete an annotation by id (removes its bookmark and custom-XML entry).</summary>
    public EditResult RemoveAnnotation(string annotationId);

    /// <summary>Mutate metadata without re-targeting. Range, id, BookmarkName are preserved.</summary>
    public EditResult UpdateAnnotation(string annotationId, AnnotationUpdate update);

    /// <summary>Atomically re-target an existing annotation to a new anchor + span.</summary>
    public EditResult MoveAnnotation(string annotationId, string newAnchorId, CharSpan? newSpan);
}
```

### Semantics

| Op | `EditResult.Modified` | `EditResult.AnnotationId` | Errors |
|----|----------------------|---------------------------|--------|
| `AddAnnotation` | The enclosing block's `Anchor` | The id used (caller-supplied or generated) | `SessionDisposed`, `AnchorNotFound`, `EmptyAnnotationSpan`, `DuplicateAnnotationId`, `OffsetOutOfRange`, `InternalError` |
| `RemoveAnnotation` | The enclosing block's `Anchor` | The id removed | `SessionDisposed`, `AnnotationNotFound`, `InternalError` |
| `UpdateAnnotation` | Empty (no markdown-visible change) | The id updated | `SessionDisposed`, `AnnotationNotFound`, `InternalError` |
| `MoveAnnotation` | Old + new enclosing blocks' `Anchor`s, deduplicated when the move stays within one block | The id moved | `SessionDisposed`, `AnnotationNotFound`, `AnchorNotFound`, `EmptyAnnotationSpan`, `OffsetOutOfRange`, `InternalError` |

`Created` and `Removed` are always empty for these ops — the bookmark + custom-XML
entry are internal, not markdown anchors. `Patch` is null (annotation ops don't
change the markdown projection; the agent's view of the document text is unchanged).

### Auto-id rule

When `annotation.Id` is null or empty on `AddAnnotation`, the session generates
`Guid.NewGuid().ToString("N").Substring(0, 16)` — a 16-char lowercase hex
string. Matches the convention used elsewhere in Docxodus for internal stable
ids. The generator retries up to 4 times on collision against existing
annotation ids before returning `DuplicateAnnotationId` (vanishingly unlikely
in practice). When the caller *does* supply an id and it collides,
`DuplicateAnnotationId` is returned immediately — no retry, since a
caller-supplied id is a deterministic key.

### Empty-span rule

`EmptyAnnotationSpan` is returned when:
- An explicit `span` has `Length == 0`.
- `span == null` (whole-block mode) and the resolved block has zero inline
  runs — i.e. an empty paragraph or a structural element with no text.

Both cases mean there is nothing to bookmark.

### Bookmark naming

`BookmarkName` is always set to `AnnotationManager.BookmarkPrefix + Id`
regardless of any value the caller supplies, matching how the existing
`AnnotationManager.AddAnnotation` behaves when the caller didn't pre-set a
bookmark name. This guarantees one annotation = one bookmark and avoids
two-annotations-sharing-a-bookmark hazards. (The
`AnnotationRange.ExistingBookmarkName` mode in the existing manager — which
attaches an annotation to a user-authored bookmark — is intentionally out of
scope for this surface.)

### `AnnotationUpdate` semantics

- `LabelId`, `Label`, `Color`, `Author`: a non-null value replaces the
  existing value; null leaves it unchanged. There is no way to clear these
  scalar fields (consistent with how every other "patch" record in the
  codebase works).
- `MetadataPatch`: a per-key merge. For each entry in the patch dict, a
  non-null value sets/replaces that key in the annotation's metadata; a null
  value removes that key. The whole patch dict being null leaves metadata
  untouched. To wipe metadata entirely, pass a `MetadataPatch` containing
  every existing key mapped to null (caller's responsibility — read first if
  they don't know the keys).

## Internal architecture

### `Docxodus/Internal/AnnotationOps.cs` (new)

```csharp
internal static class AnnotationOps
{
    public static EditResult Add(
        WordprocessingDocument doc,
        AnchorTarget anchor,
        CharSpan? span,
        DocumentAnnotation annotation);

    public static EditResult Remove(
        WordprocessingDocument doc,
        string annotationId);

    public static EditResult Update(
        WordprocessingDocument doc,
        string annotationId,
        AnnotationUpdate update);

    public static EditResult Move(
        WordprocessingDocument doc,
        string annotationId,
        AnchorTarget newAnchor,
        CharSpan? newSpan);
}
```

Responsibilities:
- Resolve the anchor's `XElement` and target part via `AnchorTarget.Resolve`.
- For `Add` and `Move` with a span: use `RunTextMap.Build(blockElement)` plus
  `ResolveRange(map, span.Start, span.Length)` to determine which runs to
  split (if span boundaries fall mid-run) and where to insert
  `<w:bookmarkStart>` and `<w:bookmarkEnd>`.
- For `Add` and `Move` with `span == null`: insert `<w:bookmarkStart>` before
  the first inline run and `<w:bookmarkEnd>` after the last inline run of the
  block.
- Allocate a fresh `w:id` for the bookmark element by scanning existing
  bookmarks in the target part (`max(existing) + 1`).
- Delegate custom-XML serialization to `AnnotationsCustomXml`.
- For `Remove` and `Move`'s un-bookmark step: locate the
  `<w:bookmarkStart>`/`<w:bookmarkEnd>` pair by name and remove them; remove
  the corresponding entry from the annotations custom-XML part.
- Compute `AnnotatedText` from the resolved span's flat text and stamp it
  onto the persisted annotation.
- Persist the modified part via `PutXDocument` before returning.

Return shape: `EditResult` with `AnnotationId` populated on success and the
appropriate `Modified` anchors (the enclosing block's `Anchor` from the
projection; `MoveAnnotation` returns both old and new block anchors).

### `Docxodus/Internal/AnnotationsCustomXml.cs` (new)

Extracts the private helpers currently in `Docxodus/AnnotationManager.cs`
(roughly lines 528–700) into a shared internal type. Public surface:

```csharp
internal static class AnnotationsCustomXml
{
    public static CustomXmlPart GetOrCreate(WordprocessingDocument doc);
    public static IReadOnlyList<DocumentAnnotation> ReadAll(WordprocessingDocument doc);
    public static DocumentAnnotation? FindById(WordprocessingDocument doc, string id);
    public static void Write(WordprocessingDocument doc, DocumentAnnotation annotation);
    public static void Remove(WordprocessingDocument doc, string annotationId);
}
```

Both the existing `AnnotationManager` (public, `WmlDocument`-shaped) and the
new `AnnotationOps` (live-session) call into this helper. The existing
`AnnotationManager.cs` is updated to delegate to `AnnotationsCustomXml`
instead of holding its own copy of the serialization logic; behavior is
preserved (the existing test suite validates this).

### `DocxSession` integration

Each new method is a thin shell, identical in shape to every other Tier C
formatting/Tier D table method:

```csharp
public EditResult AddAnnotation(
    string anchorId, CharSpan? span, DocumentAnnotation annotation)
{
    if (_disposed)
        return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
    var anchor = ResolveAnchorTarget(anchorId);
    if (anchor is null)
        return EditResult.Fail(EditErrorCode.AnchorNotFound,
            $"anchor not found: {anchorId}", anchorId);

    _history.RecordPreOp(TakeSnapshot());
    try
    {
        return AnnotationOps.Add(_doc!, anchor, span, annotation);
    }
    catch (Exception ex)
    {
        return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
    }
}
```

The undo snapshot is the existing package-byte snapshot, so `Undo()` reverses
the bookmark insertion and custom-XML write together — no annotation-specific
undo logic.

## Wire format

### `Docxodus/Internal/DocxSessionOps.cs`

Four new static methods, each ~3 lines, following the existing facade
convention (`SessionRegistry.Get(handle).Method(...)` → JSON via
`DocxSessionJson`):

```csharp
public static string AddAnnotation(int handle, string anchorId,
    CharSpan? span, string annotationJson) =>
    DocxSessionJson.Serialize(SessionRegistry.Get(handle).AddAnnotation(
        anchorId, span, DocxSessionJson.DeserializeAnnotation(annotationJson)));

public static string RemoveAnnotation(int handle, string annotationId) =>
    DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveAnnotation(annotationId));

public static string UpdateAnnotation(int handle, string annotationId, string updateJson) =>
    DocxSessionJson.Serialize(SessionRegistry.Get(handle).UpdateAnnotation(
        annotationId, DocxSessionJson.DeserializeAnnotationUpdate(updateJson)));

public static string MoveAnnotation(int handle, string annotationId,
    string newAnchorId, CharSpan? newSpan) =>
    DocxSessionJson.Serialize(SessionRegistry.Get(handle).MoveAnnotation(
        annotationId, newAnchorId, newSpan));
```

### `Docxodus/Internal/DocxSessionJson.cs`

New helpers:
- `DeserializeAnnotation(string json) → DocumentAnnotation`
- `DeserializeAnnotationUpdate(string json) → AnnotationUpdate`
- The existing serializer already handles `EditResult` (so `AnnotationId`
  serializes automatically once added).
- The existing `SerializeAnnotations` continues to serve the read side
  unchanged.

JSON shape follows the existing convention (camelCase property names,
ISO-8601 strings for dates, dictionaries as JSON objects). The
`AnnotationUpdate` deserializer must preserve explicit JSON `null` values
in `metadataPatch` — those mean "remove this key" and must NOT be conflated
with a missing key (which means "leave as-is"). Concretely: configure the
serializer for `AnnotationUpdate` with
`DefaultIgnoreCondition = JsonIgnoreCondition.Never` on the metadata patch
property so `null`-valued entries deserialize into the `IReadOnlyDictionary`
rather than being dropped.

### WASM (`wasm/DocxodusWasm/DocxSessionBridge.cs`)

Four new `[JSExport]` shells. `CharSpan?` follows the existing
`ApplyFormat` convention: a JSON string `spanJson` where empty/null means
"no span / whole block" and otherwise `{"start": int, "length": int}`.

```csharp
[JSExport]
public static string AddAnnotation(
    int h, string anchorId, string spanJson, string annotationJson)
{
    CharSpan? span = null;
    if (!string.IsNullOrEmpty(spanJson))
    {
        using var doc = JsonDocument.Parse(spanJson);
        span = new CharSpan(
            doc.RootElement.GetProperty("start").GetInt32(),
            doc.RootElement.GetProperty("length").GetInt32());
    }
    return DocxSessionOps.AddAnnotation(h, anchorId, span, annotationJson);
}
```

`UpdateAnnotation` and `RemoveAnnotation` are straight string passthroughs;
`MoveAnnotation` parses `newSpanJson` the same way.

### Python stdio host (`tools/python-host/Dispatcher.cs`)

Four new switch cases:

```csharp
"add_annotation" => DocxSessionOps.AddAnnotation(
    Handle(args), Str(args, "anchorId"), OptSpan(args, "span"),
    Str(args, "annotation")),
"remove_annotation" => DocxSessionOps.RemoveAnnotation(
    Handle(args), Str(args, "annotationId")),
"update_annotation" => DocxSessionOps.UpdateAnnotation(
    Handle(args), Str(args, "annotationId"), Str(args, "update")),
"move_annotation" => DocxSessionOps.MoveAnnotation(
    Handle(args), Str(args, "annotationId"),
    Str(args, "newAnchorId"), OptSpan(args, "newSpan")),
```

A small `OptSpan(args, name)` helper is added to parse `{ "start": int,
"length": int }` JSON objects into `CharSpan?` (null when the field is
missing or null).

### npm/TypeScript (`npm/src/`)

- `types.ts` — add `DocumentAnnotation` (already present for read side; verify
  shape covers all write-side fields) and a new `AnnotationUpdate` interface.
  Extend the `DocxodusWasmExports` interface with the four new method
  signatures.
- `index.ts` — add four methods to the `DocxSession` wrapper class. Each
  marshals `CharSpan?` to the two-int ABI and posts the JSON via the
  exported WASM method. Returns are parsed `EditResult` objects.
- `docxodus.worker.ts` / `worker-proxy.ts` — add pass-through routing entries
  so the four ops work through the off-main-thread worker proxy that the
  React layer uses.

### Python (`python/src/docx_scalpel/`)

- `types.py` — add an `AnnotationUpdate` dataclass mirroring the C# shape;
  `DocumentAnnotation` already exists for the read side.
- `errors.py` — add three new error codes (`DuplicateAnnotationId`,
  `AnnotationNotFound`, `EmptyAnnotationSpan`) to the enum/code-to-class map.
- `session.py` — add four methods to the `DocxSession` class, each a thin
  wrapper around `self._call(...)` returning a parsed `EditResult` (the
  parser auto-picks up the new `annotationId` field once it's in the JSON).

## Testing

TDD discipline: each tier's tests are written before its implementation.

### Core — `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs` (new)

Coverage:
- **Add happy path:** anchor + span, anchor + null span, caller id vs.
  auto-generated id, `AnnotatedText` populated correctly, `BookmarkName`
  generated, `Created` defaulted.
- **Add error paths:** anchor not found, session disposed, duplicate
  caller-supplied id (`DuplicateAnnotationId`), span length 0
  (`EmptyAnnotationSpan`), span outside the block (`OffsetOutOfRange`).
- **Add range mechanics:** span straddling two runs (run-split logic), span
  matching exactly one run, span at the start of the block, span at the end.
- **Remove happy + error:** by id, error when id missing.
- **Update happy + error:** scalar-only patch, metadata-only patch, mixed,
  metadata-clear-key (null value), error when id missing.
- **Move happy + error:** different block, same block new span, error when
  id missing, error when new anchor missing.
- **Undo/redo:** every op round-trips through `Undo()` + `Redo()`.
- **Interleaving:** add annotation → `ReplaceText` on the same block →
  annotation bookmark survives (Word's normal bookmark behavior under
  text edits).
- **Persistence round-trip:** add → `Save()` → reopen as new session →
  `ListAnnotations()` returns the same data.
- **Cross-part:** add an annotation in a header part (anchor in a header),
  verify it persists alongside main-document annotations.

### WASM — `npm/tests/annotations-write.spec.ts` (new)

Playwright spec running in the existing harness:
- Open fixture → `addAnnotation` → `listAnnotations` returns it.
- `updateAnnotation` → re-list verifies changes.
- `moveAnnotation` → `findByAnnotation` returns the new anchor.
- `removeAnnotation` → re-list does not return it.
- `save` → reopen via a fresh `openDocxSession` → annotations persist.
- Error paths: each EditErrorCode reachable from the wire returns a typed
  `EditResult` with `success: false` and the expected `error.code`.

### Python — `python/tests/test_annotations_write.py` (new)

Same matrix as the WASM spec but through the stdio host:
- Per-op happy path → `find_by_annotation` + `list_annotations` validation.
- Each error code is reachable and parses into the right
  `DocxScalpelError` subclass.
- Save → reopen → list round-trip preserves all annotation properties.
- One test interleaves a write annotation op with `replace_text` to confirm
  the dispatcher correctly handles ordering and the session state is
  consistent across calls.

### Smoke — extend `python/tests/test_lifecycle.py`

Add one quick assertion that an annotation written via `add_annotation` is
visible to a subsequent `list_annotations` in the same session (no reopen).

## Documentation

- **`CHANGELOG.md`** — under `[Unreleased]`:
  > Added: write surface for annotations on `DocxSession`
  > (`AddAnnotation` / `RemoveAnnotation` / `UpdateAnnotation` /
  > `MoveAnnotation`) — exposed across .NET, WASM, npm, and the Python
  > `docx-scalpel` package.

- **`CLAUDE.md`** —
  - Update the `DocxSession.cs` bullet under "Core Modules" to mention the
    new ops alongside Tier A–D.
  - In the "Feature Development Workflow" rippling reminder, add
    `python/src/docx_scalpel/session.py` to the list of layers that need
    updating when a `DocxSessionOps` method changes (it currently lists
    only the npm path).

- **`docs/architecture/docx_mutation_api.md`** — new "Tier E: Annotations"
  section covering the four methods, the `AnnotationUpdate` shape, the
  three new `EditErrorCode` values, the auto-id rule, the bookmark-name
  immutability rule, and the cross-block `MoveAnnotation` semantics.

- **`docs/architecture/python_docxodus.md`** — already up to date with the
  shipped wrapper; add the four new ops to its method inventory.

## Out of scope

- The existing `AnnotationManager.AddAnnotation(AnnotationRange)` and
  `(AnnotationTarget)` overloads remain unchanged. They cover the
  "annotate via search text / paragraph indices / table coordinates"
  flows that don't fit the anchor-addressed model.
- `ExternalAnnotationManager` (out-of-document annotation sets) is
  unchanged. The same data flow into and out of the session would be a
  separate, larger spec.
- No CLI tool changes (`tools/redline/`, `tools/docx2html/`,
  `tools/docx2oc/`).
- No web demo changes (`web/DocxodusWeb/`).
- No `AnnotationRange.ExistingBookmarkName` mode in the session surface
  (the session owns the bookmark for each annotation it creates).
