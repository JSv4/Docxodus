# Annotation Write Surface Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a write surface for Docxodus annotations to `DocxSession`, then route it through the `DocxSessionOps` facade so that the WASM bridge, the Python stdio host, the npm wrapper, and the `docx-scalpel` Python package all gain typed `AddAnnotation` / `RemoveAnnotation` / `UpdateAnnotation` / `MoveAnnotation` methods.

**Architecture:** Mutate the live `WordprocessingDocument` inside the open session. A new internal `AnnotationOps` helper handles bookmark insertion/removal at anchor + char-span boundaries; a small extracted `AnnotationsCustomXml` helper holds the persistence logic shared with the existing public `AnnotationManager`. The four `DocxSession` methods are thin shells (resolve anchor → snapshot for undo → delegate) returning the standard `EditResult` envelope.

**Tech Stack:** .NET 8 / OpenXmlSdk 3.x / xUnit / TypeScript / Playwright / Python 3.10+ / pytest.

**Spec:** `docs/superpowers/specs/2026-05-27-annotation-write-surface-design.md`

---

## File Structure

**New files:**
- `Docxodus/Internal/AnnotationsCustomXml.cs` — extracted custom-XML helpers (shared by AnnotationManager + AnnotationOps)
- `Docxodus/Internal/AnnotationOps.cs` — anchor-addressed Add/Remove/Update/Move on a live `WordprocessingDocument`
- `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs` — xUnit coverage of the four new ops
- `npm/tests/annotations-write.spec.ts` — Playwright spec for the WASM wrapper
- `python/tests/test_annotations_write.py` — pytest spec for the docx-scalpel wrapper

**Modified files (with line-anchor for the surgery point):**
- `Docxodus/AnnotationManager.cs:528-628` — delete the private custom-XML helpers, delegate to `AnnotationsCustomXml`
- `Docxodus/DocxSession.cs:579-622` — add 3 `EditErrorCode` values + `AnnotationId` field on `EditResult` + `AnnotationUpdate` record
- `Docxodus/DocxSession.cs:3000-end` — add 4 new public methods (`AddAnnotation`, `RemoveAnnotation`, `UpdateAnnotation`, `MoveAnnotation`)
- `Docxodus/Internal/DocxSessionOps.cs:86` — add 4 facade methods after "Undo / Redo" region
- `Docxodus/Internal/DocxSessionJson.cs` — add `DeserializeAnnotation` and `DeserializeAnnotationUpdate`
- `wasm/DocxodusWasm/DocxSessionBridge.cs:279` — add 4 `[JSExport]` shells after `ListAnnotations`
- `tools/python-host/Dispatcher.cs:94` — add 4 switch cases after `"list_annotations"`
- `npm/src/types.ts:1189` — extend `DocumentAnnotation` with `metadata`; add `AnnotationUpdate` interface
- `npm/src/types.ts:606`-ish (`DocxodusWasmExports`) — add 4 new method signatures
- `npm/src/index.ts` — add 4 methods on the `DocxSession` class
- `npm/src/docxodus.worker.ts` and `npm/src/worker-proxy.ts` — add 4 routing entries each
- `python/src/docx_scalpel/types.py:421-444` — extend `DocumentAnnotation` dataclass with `metadata`; add `AnnotationUpdate` dataclass
- `python/src/docx_scalpel/session.py:396` — add 4 methods after `list_annotations`
- `CHANGELOG.md` — Unreleased entry
- `CLAUDE.md` — rippling reminder + DocxSession bullet
- `docs/architecture/docx_mutation_api.md` — new Tier E section
- `docs/architecture/python_docxodus.md` — add the 4 new ops to its method inventory

---

## Phase 1: Foundation (extract custom-XML helper, add types)

### Task 1: Extract `AnnotationsCustomXml` helper

**Files:**
- Create: `Docxodus/Internal/AnnotationsCustomXml.cs`
- Modify: `Docxodus/AnnotationManager.cs` (delete private helpers, delegate)

- [ ] **Step 1: Verify existing AnnotationManager tests pass before refactor**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AnnotationManager"
```
Expected: PASS (baseline).

- [ ] **Step 2: Create `Docxodus/Internal/AnnotationsCustomXml.cs`**

```csharp
#nullable enable

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Read/write helpers for the Docxodus annotations Custom XML Part. Shared
/// between the public <see cref="AnnotationManager"/> (WmlDocument API) and
/// <see cref="AnnotationOps"/> (live DocxSession API) so the serialization
/// shape and part-discovery rules live in exactly one place.
/// </summary>
internal static class AnnotationsCustomXml
{
    public const string Namespace = AnnotationManager.AnnotationsNamespace;
    private static readonly XNamespace Ann = Namespace;

    public static CustomXmlPart? Find(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main is null) return null;
        foreach (var part in main.CustomXmlParts)
        {
            try
            {
                var xdoc = part.GetXDocument();
                if (xdoc.Root?.Name.Namespace == Ann && xdoc.Root.Name.LocalName == "annotations")
                    return part;
            }
            catch
            {
                // Not XML, or not annotations — skip.
            }
        }
        return null;
    }

    public static CustomXmlPart GetOrCreate(WordprocessingDocument doc)
    {
        var existing = Find(doc);
        if (existing is not null) return existing;

        var part = doc.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        var xdoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Ann + "annotations", new XAttribute("version", "1.0")));
        part.PutXDocument(xdoc);
        return part;
    }

    public static DocumentAnnotation? FindById(WordprocessingDocument doc, string annotationId)
    {
        var part = Find(doc);
        if (part is null) return null;
        var element = part.GetXDocument().Root?
            .Elements(Ann + "annotation")
            .FirstOrDefault(a => (string?)a.Attribute("id") == annotationId);
        return element is null ? null : Parse(element);
    }

    public static IReadOnlyList<DocumentAnnotation> ReadAll(WordprocessingDocument doc)
    {
        var part = Find(doc);
        if (part is null) return System.Array.Empty<DocumentAnnotation>();
        return part.GetXDocument().Root?
            .Elements(Ann + "annotation")
            .Select(Parse)
            .Where(a => a is not null)
            .Select(a => a!)
            .ToList() ?? new List<DocumentAnnotation>();
    }

    public static void Write(WordprocessingDocument doc, DocumentAnnotation annotation)
    {
        var part = GetOrCreate(doc);
        var xdoc = part.GetXDocument();
        var existing = xdoc.Root?
            .Elements(Ann + "annotation")
            .FirstOrDefault(a => (string?)a.Attribute("id") == annotation.Id);
        existing?.Remove();
        xdoc.Root!.Add(Serialize(annotation));
        part.PutXDocument();
    }

    public static bool Remove(WordprocessingDocument doc, string annotationId)
    {
        var part = Find(doc);
        if (part is null) return false;
        var xdoc = part.GetXDocument();
        var element = xdoc.Root?
            .Elements(Ann + "annotation")
            .FirstOrDefault(a => (string?)a.Attribute("id") == annotationId);
        if (element is null) return false;
        element.Remove();
        part.PutXDocument();
        return true;
    }

    private static XElement Serialize(DocumentAnnotation a)
    {
        var element = new XElement(Ann + "annotation",
            new XAttribute("id", a.Id),
            new XAttribute("labelId", a.LabelId ?? ""),
            new XAttribute("label", a.Label ?? ""),
            new XAttribute("color", a.Color ?? "#FFFF00"));
        if (!string.IsNullOrEmpty(a.Author)) element.Add(new XAttribute("author", a.Author));
        if (a.Created.HasValue) element.Add(new XAttribute("created", a.Created.Value.ToString("o")));
        element.Add(new XElement(Ann + "range", new XAttribute("bookmarkName", a.BookmarkName ?? "")));
        if (a.StartPage.HasValue && a.EndPage.HasValue)
        {
            element.Add(new XElement(Ann + "pageSpan",
                new XAttribute("startPage", a.StartPage.Value),
                new XAttribute("endPage", a.EndPage.Value),
                new XAttribute("stale", a.PageInfoStale ? "true" : "false")));
        }
        if (a.Metadata is { Count: > 0 })
        {
            var meta = new XElement(Ann + "metadata");
            foreach (var (key, value) in a.Metadata)
            {
                meta.Add(new XElement(Ann + "item",
                    new XAttribute("key", key),
                    value ?? ""));
            }
            element.Add(meta);
        }
        return element;
    }

    private static DocumentAnnotation? Parse(XElement element)
    {
        var id = (string?)element.Attribute("id");
        if (string.IsNullOrEmpty(id)) return null;

        var a = new DocumentAnnotation
        {
            Id = id,
            LabelId = (string?)element.Attribute("labelId") ?? "",
            Label = (string?)element.Attribute("label") ?? "",
            Color = (string?)element.Attribute("color") ?? "",
            Author = (string?)element.Attribute("author"),
        };

        var createdStr = (string?)element.Attribute("created");
        if (System.DateTime.TryParse(createdStr, out var created)) a.Created = created;

        var range = element.Element(Ann + "range");
        if (range is not null) a.BookmarkName = (string?)range.Attribute("bookmarkName");

        var pageSpan = element.Element(Ann + "pageSpan");
        if (pageSpan is not null)
        {
            if (int.TryParse((string?)pageSpan.Attribute("startPage"), out var sp)) a.StartPage = sp;
            if (int.TryParse((string?)pageSpan.Attribute("endPage"), out var ep)) a.EndPage = ep;
            a.PageInfoStale = ((string?)pageSpan.Attribute("stale"))?.ToLowerInvariant() == "true";
            if (System.DateTime.TryParse((string?)pageSpan.Attribute("computedAt"), out var ca))
                a.PageInfoComputedAt = ca;
        }

        var meta = element.Element(Ann + "metadata");
        if (meta is not null)
        {
            foreach (var item in meta.Elements(Ann + "item"))
            {
                var key = (string?)item.Attribute("key");
                if (!string.IsNullOrEmpty(key))
                    a.Metadata[key] = item.Value;
            }
        }
        return a;
    }
}
```

- [ ] **Step 3: Delegate in `Docxodus/AnnotationManager.cs`**

Replace the bodies of these private methods to call `AnnotationsCustomXml`. Open `Docxodus/AnnotationManager.cs` and:

In `AnnotationManager.GetAnnotationsInternal` (line ~421), keep the public method but make the body:

```csharp
private static List<DocumentAnnotation> GetAnnotationsInternal(WordprocessingDocument wordDoc)
{
    var annotations = Docxodus.Internal.AnnotationsCustomXml.ReadAll(wordDoc).ToList();
    foreach (var a in annotations)
        a.AnnotatedText = GetTextInBookmark(wordDoc, a.BookmarkName);
    return annotations;
}
```

In `AnnotationManager.AddAnnotationToCustomXml` (line ~528):

```csharp
private static void AddAnnotationToCustomXml(WordprocessingDocument wordDoc, DocumentAnnotation annotation)
    => Docxodus.Internal.AnnotationsCustomXml.Write(wordDoc, annotation);
```

In `AnnotationManager.RemoveAnnotationFromCustomXml` (line ~575):

```csharp
private static void RemoveAnnotationFromCustomXml(WordprocessingDocument wordDoc, string annotationId)
    => Docxodus.Internal.AnnotationsCustomXml.Remove(wordDoc, annotationId);
```

In `AnnotationManager.FindAnnotationsCustomXmlPart` (line ~591):

```csharp
private static CustomXmlPart FindAnnotationsCustomXmlPart(WordprocessingDocument wordDoc)
    => Docxodus.Internal.AnnotationsCustomXml.Find(wordDoc);
```

In `AnnotationManager.GetOrCreateAnnotationsCustomXmlPart` (line ~612):

```csharp
private static CustomXmlPart GetOrCreateAnnotationsCustomXmlPart(WordprocessingDocument wordDoc)
    => Docxodus.Internal.AnnotationsCustomXml.GetOrCreate(wordDoc);
```

Leave `ParseAnnotationElement` (line ~466) and `GetAnnotationInternal` (line ~446) in place — they're only consumed internally and refactoring them would broaden the diff.

- [ ] **Step 4: Build and rerun AnnotationManager tests**

```bash
dotnet build Docxodus.sln
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AnnotationManager"
```
Expected: PASS (no behavior change).

- [ ] **Step 5: Commit**

```bash
git add Docxodus/Internal/AnnotationsCustomXml.cs Docxodus/AnnotationManager.cs
git commit -m "refactor(annotations): extract custom-XML helpers into Internal/AnnotationsCustomXml"
```

---

### Task 2: Add `AnnotationUpdate`, error codes, `EditResult.AnnotationId`

**Files:**
- Modify: `Docxodus/DocxSession.cs` (insert at lines 579-622 region)

- [ ] **Step 1: Edit `Docxodus/DocxSession.cs` — extend `EditErrorCode`**

Find the `public enum EditErrorCode { ... }` block (line ~579) and add three values before the closing brace:

```csharp
    NothingToUndo,
    NothingToRedo,

    DuplicateAnnotationId,
    AnnotationNotFound,
    EmptyAnnotationSpan,

    InternalError,
}
```

- [ ] **Step 2: Edit `Docxodus/DocxSession.cs` — extend `EditResult`**

Find the `public sealed class EditResult` block (line ~611) and add an `AnnotationId` field:

```csharp
public sealed class EditResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public IReadOnlyList<Anchor> Created { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Removed { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Modified { get; init; } = Array.Empty<Anchor>();
    public MarkdownPatch? Patch { get; init; }

    /// <summary>
    /// Populated by AddAnnotation/RemoveAnnotation/UpdateAnnotation/MoveAnnotation
    /// with the affected annotation id. Null for every other op.
    /// </summary>
    public string? AnnotationId { get; init; }

    internal static EditResult Fail(EditErrorCode code, string message, string? anchorId = null) =>
        new() { Success = false, Error = new EditError(code, message, anchorId) };
}
```

- [ ] **Step 3: Edit `Docxodus/DocxSession.cs` — add `AnnotationUpdate` record**

Insert the new record right after `EditResult` (around line 622):

```csharp
/// <summary>
/// Partial-update payload for <see cref="DocxSession.UpdateAnnotation"/>.
/// Null fields leave the existing value unchanged. <see cref="MetadataPatch"/>
/// is a per-key merge: a non-null value sets the key, an explicit null removes
/// it, a missing key leaves it unchanged.
/// </summary>
public sealed record AnnotationUpdate
{
    public string? LabelId { get; init; }
    public string? Label { get; init; }
    public string? Color { get; init; }
    public string? Author { get; init; }
    public IReadOnlyDictionary<string, string?>? MetadataPatch { get; init; }
}
```

- [ ] **Step 4: Build to verify the new types compile**

```bash
dotnet build Docxodus/Docxodus.csproj
```
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add Docxodus/DocxSession.cs
git commit -m "feat(session): add AnnotationUpdate + EditErrorCode entries + EditResult.AnnotationId"
```

---

## Phase 2: `AddAnnotation` core

### Task 3: Failing happy-path test for `AddAnnotation`

**Files:**
- Create: `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs`

- [ ] **Step 1: Create the test file with a happy-path test**

```csharp
#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionAnnotationWriteTests
{
    private static byte[] LoadFixture(string name) =>
        File.ReadAllBytes(Path.Combine("TestFiles", name));

    // Smallest known-good fixture used throughout the suite.
    private const string Fixture = "DA001-TemplateDocument.docx";

    [Fact]
    public void AW001_AddAnnotation_ByAnchorAndSpan_PersistsBookmarkAndCustomXml()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p");

        var annotation = new DocumentAnnotation
        {
            Id = "ann-001",
            LabelId = "RISK",
            Label = "Risk",
            Color = "#FFEB3B",
            Author = "tester",
        };

        var result = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 4), annotation);

        Assert.True(result.Success);
        Assert.Equal("ann-001", result.AnnotationId);
        Assert.Single(result.Modified);
        Assert.Equal(firstP.Anchor.Id, result.Modified[0].Id);

        var listed = session.ListAnnotations();
        Assert.Single(listed, a => a.Id == "ann-001" && a.LabelId == "RISK");
    }
}
```

- [ ] **Step 2: Run test to confirm it fails**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW001_AddAnnotation_ByAnchorAndSpan"
```
Expected: FAIL with `'DocxSession' does not contain a definition for 'AddAnnotation'`.

- [ ] **Step 3: Commit the failing test**

```bash
git add Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "test(session): failing happy-path test for AddAnnotation"
```

---

### Task 4: Implement `AnnotationOps.Add` + `DocxSession.AddAnnotation` (anchor + span only)

**Files:**
- Create: `Docxodus/Internal/AnnotationOps.cs`
- Modify: `Docxodus/DocxSession.cs` (append four methods near the bottom)

- [ ] **Step 1: Create `Docxodus/Internal/AnnotationOps.cs`**

```csharp
#nullable enable

using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Anchor-addressed annotation mutations on an open <see cref="WordprocessingDocument"/>.
/// Shared backend for <see cref="DocxSession.AddAnnotation"/>,
/// <see cref="DocxSession.RemoveAnnotation"/>, <see cref="DocxSession.UpdateAnnotation"/>,
/// and <see cref="DocxSession.MoveAnnotation"/>.
/// </summary>
internal static class AnnotationOps
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static EditResult Add(
        WordprocessingDocument doc,
        AnchorTarget anchor,
        CharSpan? span,
        DocumentAnnotation annotation)
    {
        ArgumentNullException.ThrowIfNull(annotation);

        var block = anchor.Resolve(doc);
        if (block is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                "element resolved null", anchor.Anchor.Id);

        // Resolve id (auto-generate or check for collision).
        var id = string.IsNullOrEmpty(annotation.Id) ? null : annotation.Id;
        if (id is null)
        {
            id = GenerateUniqueId(doc);
            if (id is null)
                return EditResult.Fail(EditErrorCode.DuplicateAnnotationId,
                    "auto-id collided 4 times", anchor.Anchor.Id);
        }
        else if (AnnotationsCustomXml.FindById(doc, id) is not null)
        {
            return EditResult.Fail(EditErrorCode.DuplicateAnnotationId,
                $"annotation id already exists: {id}", anchor.Anchor.Id);
        }

        // Build the run text map and resolve span.
        var map = RunTextMap.Build(block);
        int spanStart, spanLength;
        if (span.HasValue)
        {
            spanStart = span.Value.Start;
            spanLength = span.Value.Length;
            if (spanLength <= 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "span length must be > 0", anchor.Anchor.Id);
            if (spanStart < 0 || spanStart + spanLength > map.FlatText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{spanStart},{spanStart + spanLength}) outside block " +
                    $"of length {map.FlatText.Length}", anchor.Anchor.Id);
        }
        else
        {
            spanStart = 0;
            spanLength = map.FlatText.Length;
            if (spanLength == 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "block has no inline runs to bookmark", anchor.Anchor.Id);
        }

        var annotatedText = map.FlatText.Substring(spanStart, spanLength);

        // Insert bookmarkStart/bookmarkEnd around the span.
        var bookmarkName = AnnotationManager.BookmarkPrefix + id;
        var bookmarkId = NextBookmarkId(block.Document!.Root!);

        var (startRunInsert, endRunInsert) = SplitRunsForSpan(map, spanStart, spanLength);

        var bookmarkStart = new XElement(W + "bookmarkStart",
            new XAttribute(W + "id", bookmarkId),
            new XAttribute(W + "name", bookmarkName));
        var bookmarkEnd = new XElement(W + "bookmarkEnd",
            new XAttribute(W + "id", bookmarkId));

        startRunInsert.AddBeforeSelf(bookmarkStart);
        endRunInsert.AddAfterSelf(bookmarkEnd);

        // Persist custom XML.
        annotation.Id = id;
        annotation.BookmarkName = bookmarkName;
        annotation.Created ??= DateTime.UtcNow;
        annotation.AnnotatedText = annotatedText;
        annotation.PageInfoStale = true;
        AnnotationsCustomXml.Write(doc, annotation);

        // Persist part XML.
        var partUri = anchor.PartUri;
        SavePart(doc, partUri);

        return new EditResult
        {
            Success = true,
            AnnotationId = id,
            Modified = new[] { anchor.Anchor },
        };
    }

    public static EditResult Remove(WordprocessingDocument doc, string annotationId)
    {
        if (string.IsNullOrEmpty(annotationId))
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                "annotation id required");

        var existing = AnnotationsCustomXml.FindById(doc, annotationId);
        if (existing is null)
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                $"annotation not found: {annotationId}");

        var bookmarkName = existing.BookmarkName;
        Anchor? touchedBlock = null;
        if (!string.IsNullOrEmpty(bookmarkName))
        {
            touchedBlock = RemoveBookmarkPair(doc, bookmarkName!);
        }

        AnnotationsCustomXml.Remove(doc, annotationId);

        return new EditResult
        {
            Success = true,
            AnnotationId = annotationId,
            Modified = touchedBlock is null
                ? Array.Empty<Anchor>()
                : new[] { touchedBlock },
        };
    }

    public static EditResult Update(
        WordprocessingDocument doc,
        string annotationId,
        AnnotationUpdate update)
    {
        ArgumentNullException.ThrowIfNull(update);

        var existing = AnnotationsCustomXml.FindById(doc, annotationId);
        if (existing is null)
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                $"annotation not found: {annotationId}");

        if (update.LabelId is not null) existing.LabelId = update.LabelId;
        if (update.Label is not null) existing.Label = update.Label;
        if (update.Color is not null) existing.Color = update.Color;
        if (update.Author is not null) existing.Author = update.Author;
        if (update.MetadataPatch is not null)
        {
            existing.Metadata ??= new System.Collections.Generic.Dictionary<string, string>();
            foreach (var (key, value) in update.MetadataPatch)
            {
                if (value is null) existing.Metadata.Remove(key);
                else existing.Metadata[key] = value;
            }
        }

        AnnotationsCustomXml.Write(doc, existing);

        return new EditResult
        {
            Success = true,
            AnnotationId = annotationId,
        };
    }

    public static EditResult Move(
        WordprocessingDocument doc,
        string annotationId,
        AnchorTarget newAnchor,
        CharSpan? newSpan)
    {
        var existing = AnnotationsCustomXml.FindById(doc, annotationId);
        if (existing is null)
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                $"annotation not found: {annotationId}");

        // Validate the new range BEFORE removing the old bookmark so we don't
        // strand the annotation.
        var newBlock = newAnchor.Resolve(doc);
        if (newBlock is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                "element resolved null", newAnchor.Anchor.Id);

        var newMap = RunTextMap.Build(newBlock);
        int s, l;
        if (newSpan.HasValue)
        {
            s = newSpan.Value.Start;
            l = newSpan.Value.Length;
            if (l <= 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "span length must be > 0", newAnchor.Anchor.Id);
            if (s < 0 || s + l > newMap.FlatText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{s},{s + l}) outside block of length {newMap.FlatText.Length}",
                    newAnchor.Anchor.Id);
        }
        else
        {
            s = 0;
            l = newMap.FlatText.Length;
            if (l == 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "block has no inline runs to bookmark", newAnchor.Anchor.Id);
        }

        var bookmarkName = existing.BookmarkName;
        Anchor? oldBlockAnchor = null;
        if (!string.IsNullOrEmpty(bookmarkName))
            oldBlockAnchor = RemoveBookmarkPair(doc, bookmarkName!);

        // Reinsert with a fresh bookmark id at the new range.
        var bookmarkId = NextBookmarkId(newBlock.Document!.Root!);
        var (startRunInsert, endRunInsert) = SplitRunsForSpan(newMap, s, l);
        startRunInsert.AddBeforeSelf(new XElement(W + "bookmarkStart",
            new XAttribute(W + "id", bookmarkId),
            new XAttribute(W + "name", bookmarkName!)));
        endRunInsert.AddAfterSelf(new XElement(W + "bookmarkEnd",
            new XAttribute(W + "id", bookmarkId)));

        existing.AnnotatedText = newMap.FlatText.Substring(s, l);
        existing.PageInfoStale = true;
        AnnotationsCustomXml.Write(doc, existing);

        SavePart(doc, newAnchor.PartUri);

        var modified = oldBlockAnchor is null || oldBlockAnchor.Id == newAnchor.Anchor.Id
            ? new[] { newAnchor.Anchor }
            : new[] { oldBlockAnchor, newAnchor.Anchor };

        return new EditResult
        {
            Success = true,
            AnnotationId = annotationId,
            Modified = modified,
        };
    }

    // ─── helpers ───────────────────────────────────────────────────────────

    private static string? GenerateUniqueId(WordprocessingDocument doc)
    {
        for (int i = 0; i < 4; i++)
        {
            var candidate = Guid.NewGuid().ToString("N").Substring(0, 16);
            if (AnnotationsCustomXml.FindById(doc, candidate) is null)
                return candidate;
        }
        return null;
    }

    private static int NextBookmarkId(XElement root)
    {
        var max = root.Descendants(W + "bookmarkStart")
            .Select(b => (int?)b.Attribute(W + "id"))
            .Where(v => v.HasValue)
            .Select(v => v!.Value)
            .DefaultIfEmpty(0)
            .Max();
        return max + 1;
    }

    /// <summary>
    /// Splits the runs at the span boundaries (when boundaries fall mid-run) so
    /// that <c>w:bookmarkStart</c> can be inserted before the run containing the
    /// span start and <c>w:bookmarkEnd</c> after the run containing the span end,
    /// with no other runs between them inside the span.
    /// Returns the start-side and end-side runs to insert before/after.
    /// </summary>
    private static (XElement startRun, XElement endRun) SplitRunsForSpan(
        RunTextMap.Map map, int start, int length)
    {
        var segments = RunTextMap.ResolveRange(map, start, length);
        // segments is guaranteed non-empty when length > 0 and bounds were checked.

        var first = segments[0];
        var last = segments[^1];

        // Split the first run if the span starts mid-run.
        XElement startRun = first.Segment.Run;
        if (first.OffsetInRun > 0)
        {
            startRun = SplitRunAt(startRun, first.OffsetInRun, takeRightHalf: true);
        }

        // Split the last run if the span ends mid-run.
        XElement endRun = last.Segment.Run;
        if (last.OffsetInRun + last.Length < last.Segment.Length)
        {
            endRun = SplitRunAt(endRun, last.OffsetInRun + last.Length, takeRightHalf: false);
        }

        // When the start and end were originally the same run AND we split off
        // a right-half above for startRun, the endRun reference must follow that.
        if (first.Segment.Run == last.Segment.Run && startRun != first.Segment.Run)
            endRun = startRun;

        return (startRun, endRun);
    }

    /// <summary>
    /// Splits a <c>w:r</c> element at <paramref name="offsetInRunText"/> characters
    /// from the start of its text. Returns either the original (left) run or the
    /// newly-inserted right-hand run, per <paramref name="takeRightHalf"/>.
    /// Only handles the common case of a run with a single <c>w:t</c> child;
    /// runs with multiple text fragments or other content are out of scope for
    /// v1 (the bookmark just sits at the closer of the two boundary positions).
    /// </summary>
    private static XElement SplitRunAt(XElement run, int offsetInRunText, bool takeRightHalf)
    {
        var text = run.Element(W + "t");
        if (text is null) return run;

        var full = text.Value;
        if (offsetInRunText <= 0 || offsetInRunText >= full.Length) return run;

        var leftText = full.Substring(0, offsetInRunText);
        var rightText = full.Substring(offsetInRunText);

        text.Value = leftText;
        if (string.IsNullOrEmpty(text.Attribute(XNamespace.Xml + "space")?.Value)
            && (leftText.StartsWith(' ') || leftText.EndsWith(' ')))
        {
            text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        }

        var rightRun = new XElement(run); // clone formatting
        var rightText_e = rightRun.Element(W + "t")!;
        rightText_e.Value = rightText;
        if (rightText.StartsWith(' ') || rightText.EndsWith(' '))
            rightText_e.SetAttributeValue(XNamespace.Xml + "space", "preserve");

        run.AddAfterSelf(rightRun);
        return takeRightHalf ? rightRun : run;
    }

    private static Anchor? RemoveBookmarkPair(WordprocessingDocument doc, string bookmarkName)
    {
        Anchor? affectedBlock = null;
        foreach (var part in EnumerateParts(doc))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            var start = root.Descendants(W + "bookmarkStart")
                .FirstOrDefault(b => (string?)b.Attribute(W + "name") == bookmarkName);
            if (start is null) continue;

            var id = (string?)start.Attribute(W + "id");
            var end = id is null
                ? null
                : root.Descendants(W + "bookmarkEnd")
                    .FirstOrDefault(b => (string?)b.Attribute(W + "id") == id);

            // Locate enclosing block for the Modified anchor.
            var enclosing = start.AncestorsAndSelf()
                .FirstOrDefault(e => (string?)e.Attribute(PtOpenXml.Unid) is not null);
            if (enclosing is not null)
            {
                affectedBlock = new Anchor(
                    Kind: KindOfElement(enclosing),
                    Scope: ScopeOfPart(part),
                    Unid: (string)enclosing.Attribute(PtOpenXml.Unid)!,
                    Id: $"{KindOfElement(enclosing)}:{ScopeOfPart(part)}:{(string)enclosing.Attribute(PtOpenXml.Unid)!}");
            }

            start.Remove();
            end?.Remove();
            part.PutXDocument();
            break;
        }
        return affectedBlock;
    }

    private static System.Collections.Generic.IEnumerable<OpenXmlPart> EnumerateParts(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }

    private static void SavePart(WordprocessingDocument doc, string partUri)
    {
        foreach (var part in EnumerateParts(doc))
        {
            if (part.Uri.ToString() == partUri)
            {
                part.PutXDocument();
                return;
            }
        }
    }

    private static string KindOfElement(XElement e)
    {
        if (e.Name == W + "p")
        {
            var pStyle = e.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            if (pStyle?.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) == true) return "h";
            if (e.Element(W + "pPr")?.Element(W + "numPr") is not null) return "li";
            return "p";
        }
        if (e.Name == W + "tbl") return "tbl";
        if (e.Name == W + "tr") return "tr";
        if (e.Name == W + "tc") return "tc";
        return "p";
    }

    private static string ScopeOfPart(OpenXmlPart part)
    {
        if (part is MainDocumentPart) return "body";
        if (part is HeaderPart) return "hdr";
        if (part is FooterPart) return "ftr";
        if (part is FootnotesPart) return "fn";
        if (part is EndnotesPart) return "en";
        return "body";
    }
}
```

- [ ] **Step 2: Add `DocxSession.AddAnnotation` and three other thin shells**

Append at the end of `Docxodus/DocxSession.cs` (after the last existing method, before the closing brace of the `DocxSession` class):

```csharp
    // ─── Tier E: annotations ────────────────────────────────────────────

    /// <summary>
    /// Annotate the range <paramref name="span"/> inside the block addressed by
    /// <paramref name="anchorId"/>. When <paramref name="span"/> is null, the
    /// annotation wraps every inline run of the block. When
    /// <paramref name="annotation"/>.Id is null/empty, a 16-char hex id is
    /// generated. The bookmark name, AnnotatedText, Created, and PageInfoStale
    /// fields of the annotation are always set by this method.
    /// </summary>
    public EditResult AddAnnotation(string anchorId, CharSpan? span, DocumentAnnotation annotation)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (annotation is null)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "annotation is null", anchorId);

        var anchor = ResolveAnchorTarget(anchorId);
        if (anchor is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            return Docxodus.Internal.AnnotationOps.Add(_doc!, anchor, span, annotation);
        }
        catch (Exception ex)
        {
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>Removes an annotation (its bookmark and custom-XML entry) by id.</summary>
    public EditResult RemoveAnnotation(string annotationId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            return Docxodus.Internal.AnnotationOps.Remove(_doc!, annotationId);
        }
        catch (Exception ex)
        {
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>Mutates label/color/author/metadata of an annotation without re-targeting.</summary>
    public EditResult UpdateAnnotation(string annotationId, AnnotationUpdate update)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (update is null)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "update is null");
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            return Docxodus.Internal.AnnotationOps.Update(_doc!, annotationId, update);
        }
        catch (Exception ex)
        {
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>Re-targets an existing annotation to a new anchor + span.</summary>
    public EditResult MoveAnnotation(string annotationId, string newAnchorId, CharSpan? newSpan)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var anchor = ResolveAnchorTarget(newAnchorId);
        if (anchor is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"anchor not found: {newAnchorId}", newAnchorId);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            return Docxodus.Internal.AnnotationOps.Move(_doc!, annotationId, anchor, newSpan);
        }
        catch (Exception ex)
        {
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, newAnchorId);
        }
    }
```

`ResolveAnchorTarget` is the existing private helper used by other tiers; if its current visibility is `private static`, leave it as-is — these new methods live in the same class.

- [ ] **Step 3: Build**

```bash
dotnet build Docxodus/Docxodus.csproj
```
Expected: PASS.

- [ ] **Step 4: Run the AW001 test**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW001_AddAnnotation_ByAnchorAndSpan"
```
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add Docxodus/Internal/AnnotationOps.cs Docxodus/DocxSession.cs
git commit -m "feat(session): add AddAnnotation/RemoveAnnotation/UpdateAnnotation/MoveAnnotation"
```

---

### Task 5: Add tests + behavior for null-span = whole block, error paths, edge cases

**Files:**
- Modify: `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs`

- [ ] **Step 1: Add tests for null-span, error paths, multi-run span, and round-trips**

Append these tests to `DocxSessionAnnotationWriteTests.cs`:

```csharp
    [Fact]
    public void AW002_AddAnnotation_NullSpan_BookmarksWholeBlock()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p" && a.TextPreview.Length > 0);

        var result = session.AddAnnotation(firstP.Anchor.Id, span: null,
            new DocumentAnnotation { Id = "ann-whole", LabelId = "L", Label = "L", Color = "#FFF" });

        Assert.True(result.Success);
        var listed = session.ListAnnotations().Single(a => a.Id == "ann-whole");
        Assert.Equal(firstP.TextPreview, listed.AnnotatedText);
    }

    [Fact]
    public void AW003_AddAnnotation_AnchorNotFound_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var result = session.AddAnnotation("p:body:DEADBEEFDEADBEEF", new CharSpan(0, 1),
            new DocumentAnnotation { Id = "x", LabelId = "L", Label = "L", Color = "#000" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, result.Error!.Code);
    }

    [Fact]
    public void AW004_AddAnnotation_EmptySpan_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        var result = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 0),
            new DocumentAnnotation { Id = "x", LabelId = "L", Label = "L", Color = "#000" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.EmptyAnnotationSpan, result.Error!.Code);
    }

    [Fact]
    public void AW005_AddAnnotation_OutOfRangeSpan_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p" && a.TextPreview.Length > 0);

        var result = session.AddAnnotation(firstP.Anchor.Id,
            new CharSpan(0, firstP.TextPreview.Length + 9999),
            new DocumentAnnotation { Id = "x", LabelId = "L", Label = "L", Color = "#000" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, result.Error!.Code);
    }

    [Fact]
    public void AW006_AddAnnotation_DuplicateCallerId_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        var first = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "dup", LabelId = "L", Label = "L", Color = "#000" });
        Assert.True(first.Success);

        var second = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "dup", LabelId = "L", Label = "L", Color = "#000" });
        Assert.False(second.Success);
        Assert.Equal(EditErrorCode.DuplicateAnnotationId, second.Error!.Code);
    }

    [Fact]
    public void AW007_AddAnnotation_AutoId_GeneratesUnique16Hex()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        var result = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { LabelId = "L", Label = "L", Color = "#000" });

        Assert.True(result.Success);
        Assert.NotNull(result.AnnotationId);
        Assert.Equal(16, result.AnnotationId!.Length);
        Assert.Matches("^[0-9a-f]{16}$", result.AnnotationId);
    }

    [Fact]
    public void AW008_AddAnnotation_UndoRollsBack()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "undoable", LabelId = "L", Label = "L", Color = "#000" });
        Assert.Single(session.ListAnnotations(), a => a.Id == "undoable");

        Assert.True(session.Undo());
        Assert.Empty(session.ListAnnotations().Where(a => a.Id == "undoable"));

        Assert.True(session.Redo());
        Assert.Single(session.ListAnnotations(), a => a.Id == "undoable");
    }

    [Fact]
    public void AW009_AddAnnotation_SaveAndReopen_Persists()
    {
        byte[] saved;
        using (var session = new DocxSession(LoadFixture(Fixture)))
        {
            var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
            var r = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 4),
                new DocumentAnnotation { Id = "persist", LabelId = "PERSIST", Label = "P", Color = "#0F0",
                    Metadata = new System.Collections.Generic.Dictionary<string, string> { ["k"] = "v" } });
            Assert.True(r.Success);
            saved = session.Save();
        }

        using var reopened = new DocxSession(saved);
        var found = reopened.ListAnnotations().Single(a => a.Id == "persist");
        Assert.Equal("PERSIST", found.LabelId);
        Assert.Equal("v", found.Metadata["k"]);
    }
```

- [ ] **Step 2: Run the new tests**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionAnnotationWriteTests"
```
Expected: AW001–AW009 PASS.

- [ ] **Step 3: Commit**

```bash
git add Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "test(session): cover AddAnnotation error paths, undo/redo, persistence"
```

---

### Task 6: Multi-run span + cross-part (header) tests

**Files:**
- Modify: `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs`

- [ ] **Step 1: Add the two remaining Add-coverage tests**

Append:

```csharp
    [Fact]
    public void AW010_AddAnnotation_SpanStraddlingTwoRuns_SplitsRunsCorrectly()
    {
        // Use a fixture that has multi-run paragraphs (bold + plain in same para).
        // DA001-TemplateDocument.docx has a paragraph with a content control;
        // any multi-run paragraph works. We pick the first whose flat text is at
        // least 10 characters AND has more than one inline run.
        using var session = new DocxSession(LoadFixture(Fixture));
        var target = session.AnchorsByScope(ProjectionScopes.Body)
            .Where(a => a.Anchor.Kind == "p" && a.TextPreview.Length >= 6)
            .First();

        var r = session.AddAnnotation(target.Anchor.Id, new CharSpan(2, 4),
            new DocumentAnnotation { Id = "mid", LabelId = "L", Label = "L", Color = "#000" });
        Assert.True(r.Success);

        var saved = session.Save();
        using var reopened = new DocxSession(saved);
        var listed = reopened.ListAnnotations().Single(a => a.Id == "mid");
        Assert.Equal(4, listed.AnnotatedText!.Length);
    }

    [Fact]
    public void AW011_AddAnnotation_InHeaderPart_Persists()
    {
        // Find a fixture with a header that has at least one paragraph.
        // HC003 (HtmlConverter test) has header/footer content.
        var fixturesWithHeaders = new[] { "HC003-Headers-And-Footers.docx", Fixture };
        var picked = fixturesWithHeaders.FirstOrDefault(File.Exists)
            ?? throw new Xunit.Sdk.XunitException("no header fixture available");

        using var session = new DocxSession(LoadFixture(picked));
        var headerAnchor = session.AnchorsByScope(ProjectionScopes.Headers).FirstOrDefault();
        if (headerAnchor is null) return; // skip when no headers present

        var r = session.AddAnnotation(headerAnchor.Anchor.Id, span: null,
            new DocumentAnnotation { Id = "hdr-ann", LabelId = "H", Label = "H", Color = "#0FF" });
        Assert.True(r.Success);
        Assert.Single(session.ListAnnotations(), a => a.Id == "hdr-ann");
    }
```

- [ ] **Step 2: Run them**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW010 | FullyQualifiedName~AW011"
```
Expected: PASS. (AW011 may early-return when the chosen fixture has no headers — that's intentional.)

- [ ] **Step 3: Commit**

```bash
git add Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "test(session): cover multi-run span splits + header-part annotations"
```

---

## Phase 3: `RemoveAnnotation`

### Task 7: Tests + behavior verification for RemoveAnnotation

**Files:**
- Modify: `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs`

- [ ] **Step 1: Add Remove tests**

Append:

```csharp
    [Fact]
    public void AW020_RemoveAnnotation_HappyPath_DropsBookmarkAndCustomXml()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
        session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "to-remove", LabelId = "L", Label = "L", Color = "#000" });

        var r = session.RemoveAnnotation("to-remove");
        Assert.True(r.Success);
        Assert.Equal("to-remove", r.AnnotationId);
        Assert.Empty(session.ListAnnotations().Where(a => a.Id == "to-remove"));
    }

    [Fact]
    public void AW021_RemoveAnnotation_Missing_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var r = session.RemoveAnnotation("does-not-exist");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnnotationNotFound, r.Error!.Code);
    }

    [Fact]
    public void AW022_RemoveAnnotation_UndoRestores()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
        session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "undoable-rm", LabelId = "L", Label = "L", Color = "#000" });

        session.RemoveAnnotation("undoable-rm");
        Assert.Empty(session.ListAnnotations().Where(a => a.Id == "undoable-rm"));

        Assert.True(session.Undo());
        Assert.Single(session.ListAnnotations(), a => a.Id == "undoable-rm");
    }
```

- [ ] **Step 2: Run them**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW02"
```
Expected: PASS (Remove logic was already implemented in Task 4).

- [ ] **Step 3: Commit**

```bash
git add Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "test(session): cover RemoveAnnotation happy + error + undo paths"
```

---

## Phase 4: `UpdateAnnotation`

### Task 8: Tests + behavior verification for UpdateAnnotation

**Files:**
- Modify: `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs`

- [ ] **Step 1: Add Update tests**

Append:

```csharp
    [Fact]
    public void AW030_UpdateAnnotation_ScalarPatch_MutatesFields()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
        session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "u1", LabelId = "OLD", Label = "Old", Color = "#000",
                Author = "alice" });

        var r = session.UpdateAnnotation("u1", new AnnotationUpdate
        {
            Label = "New",
            Color = "#FFF",
            Author = "bob",
        });
        Assert.True(r.Success);

        var listed = session.ListAnnotations().Single(a => a.Id == "u1");
        Assert.Equal("New", listed.Label);
        Assert.Equal("#FFF", listed.Color);
        Assert.Equal("bob", listed.Author);
        Assert.Equal("OLD", listed.LabelId); // unchanged
    }

    [Fact]
    public void AW031_UpdateAnnotation_MetadataPatch_AddsAndRemovesKeys()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
        session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation
            {
                Id = "u2", LabelId = "L", Label = "L", Color = "#000",
                Metadata = new System.Collections.Generic.Dictionary<string, string>
                {
                    ["keep"] = "yes",
                    ["drop"] = "old",
                },
            });

        var r = session.UpdateAnnotation("u2", new AnnotationUpdate
        {
            MetadataPatch = new System.Collections.Generic.Dictionary<string, string?>
            {
                ["drop"] = null,        // remove
                ["new"] = "fresh",      // add
            },
        });
        Assert.True(r.Success);

        var listed = session.ListAnnotations().Single(a => a.Id == "u2");
        Assert.Equal("yes", listed.Metadata["keep"]);
        Assert.False(listed.Metadata.ContainsKey("drop"));
        Assert.Equal("fresh", listed.Metadata["new"]);
    }

    [Fact]
    public void AW032_UpdateAnnotation_Missing_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var r = session.UpdateAnnotation("nope", new AnnotationUpdate { Label = "x" });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnnotationNotFound, r.Error!.Code);
    }
```

- [ ] **Step 2: Run them**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW03"
```
Expected: PASS.

- [ ] **Step 3: Commit**

```bash
git add Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "test(session): cover UpdateAnnotation scalar + metadata patch paths"
```

---

## Phase 5: `MoveAnnotation`

### Task 9: Tests + behavior verification for MoveAnnotation

**Files:**
- Modify: `Docxodus.Tests/DocxSessionAnnotationWriteTests.cs`

- [ ] **Step 1: Add Move tests**

Append:

```csharp
    [Fact]
    public void AW040_MoveAnnotation_DifferentBlock_ReturnsOldAndNewModified()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var paragraphs = session.AnchorsByScope(ProjectionScopes.Body)
            .Where(a => a.Anchor.Kind == "p" && a.TextPreview.Length > 0)
            .Take(2)
            .ToArray();
        Assert.Equal(2, paragraphs.Length);

        session.AddAnnotation(paragraphs[0].Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "movable", LabelId = "L", Label = "L", Color = "#000" });

        var r = session.MoveAnnotation("movable", paragraphs[1].Anchor.Id, new CharSpan(0, 2));
        Assert.True(r.Success);
        Assert.Equal("movable", r.AnnotationId);
        Assert.Equal(2, r.Modified.Count);
        Assert.Contains(r.Modified, m => m.Id == paragraphs[0].Anchor.Id);
        Assert.Contains(r.Modified, m => m.Id == paragraphs[1].Anchor.Id);
    }

    [Fact]
    public void AW041_MoveAnnotation_SameBlockNewSpan_DedupsModified()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var p = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p" && a.TextPreview.Length >= 5);

        session.AddAnnotation(p.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "shift", LabelId = "L", Label = "L", Color = "#000" });

        var r = session.MoveAnnotation("shift", p.Anchor.Id, new CharSpan(2, 2));
        Assert.True(r.Success);
        Assert.Single(r.Modified);
        Assert.Equal(p.Anchor.Id, r.Modified[0].Id);
    }

    [Fact]
    public void AW042_MoveAnnotation_AnnotationMissing_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var p = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
        var r = session.MoveAnnotation("nope", p.Anchor.Id, null);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnnotationNotFound, r.Error!.Code);
    }

    [Fact]
    public void AW043_MoveAnnotation_NewAnchorMissing_ReturnsErrorWithoutDamagingOld()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var p = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
        session.AddAnnotation(p.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "safe", LabelId = "L", Label = "L", Color = "#000" });

        var r = session.MoveAnnotation("safe", "p:body:DEADBEEFDEADBEEF", new CharSpan(0, 1));
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, r.Error!.Code);
        Assert.Single(session.ListAnnotations(), a => a.Id == "safe"); // still present
    }
```

- [ ] **Step 2: Run them**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW04"
```
Expected: PASS.

- [ ] **Step 3: Run the full new-tests file**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionAnnotationWriteTests"
```
Expected: all AW0* tests PASS.

- [ ] **Step 4: Commit**

```bash
git add Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "test(session): cover MoveAnnotation cross-block, same-block, and error paths"
```

---

## Phase 6: Wire format (`DocxSessionOps` + JSON)

### Task 10: Add JSON (de)serializers

**Files:**
- Modify: `Docxodus/Internal/DocxSessionJson.cs`

- [ ] **Step 1: Inspect the existing serializer to confirm where to add helpers**

```bash
grep -n "DeserializeFormatOp\|ParseFormatOp\|SerializeAnnotations" Docxodus/Internal/DocxSessionJson.cs | head
```
Expected: shows existing parse/serialize methods you'll mirror.

- [ ] **Step 2: Add deserializers**

Open `Docxodus/Internal/DocxSessionJson.cs` and add (at the end of the class, before the closing brace):

```csharp
    private static readonly System.Text.Json.JsonSerializerOptions s_annotationOptions =
        new(System.Text.Json.JsonSerializerDefaults.Web)
        {
            // Keep explicit `null` values in metadata patches so they round-trip as
            // "remove this key" rather than "key not present".
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.Never,
        };

    public static DocumentAnnotation DeserializeAnnotation(string json)
    {
        var a = System.Text.Json.JsonSerializer.Deserialize<DocumentAnnotation>(
            json, s_annotationOptions)
            ?? throw new System.ArgumentException("annotation JSON deserialized to null");
        // System.Text.Json constructs an empty Metadata dict; ensure it's non-null
        // even if the caller didn't include the field.
        a.Metadata ??= new System.Collections.Generic.Dictionary<string, string>();
        return a;
    }

    public static AnnotationUpdate DeserializeAnnotationUpdate(string json) =>
        System.Text.Json.JsonSerializer.Deserialize<AnnotationUpdate>(json, s_annotationOptions)
            ?? throw new System.ArgumentException("annotation update JSON deserialized to null");
```

- [ ] **Step 3: Build**

```bash
dotnet build Docxodus/Docxodus.csproj
```
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/Internal/DocxSessionJson.cs
git commit -m "feat(json): add DeserializeAnnotation + DeserializeAnnotationUpdate"
```

---

### Task 11: Add `DocxSessionOps` facade methods

**Files:**
- Modify: `Docxodus/Internal/DocxSessionOps.cs`

- [ ] **Step 1: Add four facade methods**

Open `Docxodus/Internal/DocxSessionOps.cs`. Just before the `// ─── Undo / Redo` region (around line 186), insert:

```csharp
    // ─── Tier E: annotations ────────────────────────────────────────────

    public static string AddAnnotation(int handle, string anchorId, CharSpan? span,
        string annotationJson) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).AddAnnotation(
            anchorId, span, DocxSessionJson.DeserializeAnnotation(annotationJson)));

    public static string RemoveAnnotation(int handle, string annotationId) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).RemoveAnnotation(annotationId));

    public static string UpdateAnnotation(int handle, string annotationId, string updateJson) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).UpdateAnnotation(
            annotationId, DocxSessionJson.DeserializeAnnotationUpdate(updateJson)));

    public static string MoveAnnotation(int handle, string annotationId, string newAnchorId,
        CharSpan? newSpan) =>
        DocxSessionJson.Serialize(SessionRegistry.Get(handle).MoveAnnotation(
            annotationId, newAnchorId, newSpan));
```

- [ ] **Step 2: Build**

```bash
dotnet build Docxodus.sln
```
Expected: PASS.

- [ ] **Step 3: Smoke test through `DocxSessionOps` directly**

Add this test at the bottom of `DocxSessionAnnotationWriteTests.cs`:

```csharp
    [Fact]
    public void AW050_DocxSessionOps_AddAnnotation_JsonRoundtrip()
    {
        var handle = Docxodus.Internal.SessionRegistry.OpenSession(LoadFixture(Fixture), null);
        try
        {
            var session = Docxodus.Internal.SessionRegistry.Get(handle);
            var anchorId = session.AnchorsByScope(ProjectionScopes.Body)
                .First(a => a.Anchor.Kind == "p" && a.TextPreview.Length > 0).Anchor.Id;
            var annJson = "{\"id\":\"json-001\",\"labelId\":\"X\",\"label\":\"X\",\"color\":\"#0F0\"}";
            var resultJson = Docxodus.Internal.DocxSessionOps.AddAnnotation(
                handle, anchorId, new CharSpan(0, 1), annJson);
            Assert.Contains("\"success\":true", resultJson);
            Assert.Contains("\"annotationId\":\"json-001\"", resultJson);
        }
        finally
        {
            Docxodus.Internal.SessionRegistry.CloseSession(handle);
        }
    }
```

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~AW050"`
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/Internal/DocxSessionOps.cs Docxodus.Tests/DocxSessionAnnotationWriteTests.cs
git commit -m "feat(ops): expose annotation write ops on DocxSessionOps facade"
```

---

## Phase 7: WASM bridge + npm wrapper

### Task 12: Add `[JSExport]` shells

**Files:**
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs`

- [ ] **Step 1: Add four JSExports after the existing `ListAnnotations`**

Open `wasm/DocxodusWasm/DocxSessionBridge.cs`. After the `ListAnnotations` method (line ~279), add:

```csharp
    // ─── Tier E: annotations (write surface) ──────────────────────────────

    /// <summary>
    /// Bridge for <see cref="DocxSession.AddAnnotation"/>. The span is encoded as
    /// a JSON string (empty/null = no span = annotate whole block, otherwise
    /// <c>{"start": int, "length": int}</c>) matching the existing
    /// <see cref="ApplyFormat"/> convention. The annotation JSON is a camelCase
    /// mirror of <see cref="DocumentAnnotation"/>.
    /// </summary>
    [JSExport]
    public static string AddAnnotation(int h, string anchorId, string spanJson, string annotationJson) =>
        DocxSessionOps.AddAnnotation(h, anchorId, ParseSpan(spanJson), annotationJson);

    [JSExport]
    public static string RemoveAnnotation(int h, string annotationId) =>
        DocxSessionOps.RemoveAnnotation(h, annotationId);

    [JSExport]
    public static string UpdateAnnotation(int h, string annotationId, string updateJson) =>
        DocxSessionOps.UpdateAnnotation(h, annotationId, updateJson);

    [JSExport]
    public static string MoveAnnotation(int h, string annotationId, string newAnchorId, string newSpanJson) =>
        DocxSessionOps.MoveAnnotation(h, annotationId, newAnchorId, ParseSpan(newSpanJson));

    private static CharSpan? ParseSpan(string json)
    {
        if (string.IsNullOrEmpty(json)) return null;
        using var doc = JsonDocument.Parse(json);
        return new CharSpan(
            doc.RootElement.GetProperty("start").GetInt32(),
            doc.RootElement.GetProperty("length").GetInt32());
    }
```

If `JsonDocument` isn't already in the using list at the top, add `using System.Text.Json;`.

- [ ] **Step 2: Build the WASM target**

```bash
./scripts/build-wasm.sh
```
Expected: build succeeds (output indicates WASM artifact emitted to `Docxodus/bin/Release/...`).

- [ ] **Step 3: Commit**

```bash
git add wasm/DocxodusWasm/DocxSessionBridge.cs
git commit -m "feat(wasm): expose annotation write ops via JSExport"
```

---

### Task 13: Extend npm types

**Files:**
- Modify: `npm/src/types.ts`

- [ ] **Step 1: Extend `DocumentAnnotation`**

Open `npm/src/types.ts`. Find `export interface DocumentAnnotation` (line ~1189) and add a `metadata` field:

```typescript
export interface DocumentAnnotation {
  id: string;
  labelId: string;
  label: string;
  color: string;
  bookmarkName: string;
  author?: string;
  created?: string;
  annotatedText?: string;
  /** Arbitrary string→string metadata bag persisted with the annotation. */
  metadata?: Record<string, string>;
}
```

- [ ] **Step 2: Add `AnnotationUpdate`**

Add right after `DocumentAnnotation`:

```typescript
/**
 * Partial-update payload for {@link DocxSession.updateAnnotation}.
 * Null/missing fields leave the existing value unchanged. `metadataPatch`
 * is a per-key merge: a non-null value sets the key, an explicit `null`
 * removes it, a missing key leaves it unchanged.
 */
export interface AnnotationUpdate {
  labelId?: string;
  label?: string;
  color?: string;
  author?: string;
  metadataPatch?: Record<string, string | null>;
}
```

- [ ] **Step 3: Extend `DocxodusWasmExports`**

Find the `DocxodusWasmExports` interface (line ~606 has the existing `RemoveAnnotation` for the WmlDocument API — leave that). Locate the *session* exports section (where `FindByAnnotation` lives, line ~763) and add four new entries:

```typescript
    // Session annotation write surface
    AddAnnotation: (
      handle: number,
      anchorId: string,
      spanJson: string,
      annotationJson: string
    ) => string;
    UpdateAnnotation: (
      handle: number,
      annotationId: string,
      updateJson: string
    ) => string;
    MoveAnnotation: (
      handle: number,
      annotationId: string,
      newAnchorId: string,
      newSpanJson: string
    ) => string;
    // Session RemoveAnnotation: the existing WmlDocument-style one stays;
    // the session-style one is named SessionRemoveAnnotation to disambiguate.
    SessionRemoveAnnotation: (handle: number, annotationId: string) => string;
```

> **Note:** the existing `RemoveAnnotation` at types.ts:606 is the byte-array WmlDocument API; the session-style remove is a different signature. The simplest disambiguation is the `Session` prefix on the new entry. In the JSExport in Task 12, rename `RemoveAnnotation` → `SessionRemoveAnnotation` to match. Update Task 12's JSExport accordingly before continuing.

- [ ] **Step 4: Update the Task 12 JSExport name**

Edit `wasm/DocxodusWasm/DocxSessionBridge.cs` to rename `RemoveAnnotation` to `SessionRemoveAnnotation`:

```csharp
    [JSExport]
    public static string SessionRemoveAnnotation(int h, string annotationId) =>
        DocxSessionOps.RemoveAnnotation(h, annotationId);
```

Rebuild WASM:

```bash
./scripts/build-wasm.sh
```

- [ ] **Step 5: TypeScript compile check**

```bash
cd npm && npx tsc --noEmit
```
Expected: PASS (no type errors).

- [ ] **Step 6: Commit**

```bash
git add npm/src/types.ts wasm/DocxodusWasm/DocxSessionBridge.cs
git commit -m "feat(npm): add DocumentAnnotation.metadata + AnnotationUpdate + session write exports"
```

---

### Task 14: Add wrapper methods on the npm `DocxSession` class

**Files:**
- Modify: `npm/src/index.ts`

- [ ] **Step 1: Find the session-wrapper class and existing `listAnnotations`**

```bash
grep -n "listAnnotations\|findByAnnotation" npm/src/index.ts | head
```

- [ ] **Step 2: Add four methods on the class**

Right after the existing read methods (e.g. after `listAnnotations`), add:

```typescript
  /**
   * Annotate a range inside `anchorId`. When `span` is `null`/`undefined`
   * the annotation wraps every inline run of the block. When
   * `annotation.id` is `undefined`, a 16-char hex id is auto-generated and
   * returned in `EditResult.annotationId`.
   */
  addAnnotation(
    anchorId: string,
    span: CharSpan | null,
    annotation: DocumentAnnotation,
  ): EditResult {
    const spanJson = span ? JSON.stringify(span) : "";
    const json = this._exports.AddAnnotation(
      this._handle,
      anchorId,
      spanJson,
      JSON.stringify(annotation),
    );
    return JSON.parse(json) as EditResult;
  }

  removeAnnotation(annotationId: string): EditResult {
    const json = this._exports.SessionRemoveAnnotation(this._handle, annotationId);
    return JSON.parse(json) as EditResult;
  }

  updateAnnotation(
    annotationId: string,
    update: AnnotationUpdate,
  ): EditResult {
    const json = this._exports.UpdateAnnotation(
      this._handle,
      annotationId,
      JSON.stringify(update),
    );
    return JSON.parse(json) as EditResult;
  }

  moveAnnotation(
    annotationId: string,
    newAnchorId: string,
    newSpan: CharSpan | null,
  ): EditResult {
    const spanJson = newSpan ? JSON.stringify(newSpan) : "";
    const json = this._exports.MoveAnnotation(
      this._handle,
      annotationId,
      newAnchorId,
      spanJson,
    );
    return JSON.parse(json) as EditResult;
  }
```

If `AnnotationUpdate` isn't already imported at the top of the file, add it to the import list from `./types`.

- [ ] **Step 3: TypeScript compile check**

```bash
cd npm && npx tsc --noEmit
```
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add npm/src/index.ts
git commit -m "feat(npm): add DocxSession.addAnnotation/removeAnnotation/updateAnnotation/moveAnnotation"
```

---

### Task 15: Wire the worker proxy

**Files:**
- Modify: `npm/src/docxodus.worker.ts`
- Modify: `npm/src/worker-proxy.ts`

- [ ] **Step 1: Inspect the worker routing pattern**

```bash
grep -n "addAnnotation\|listAnnotations\|replaceText\|case \"" npm/src/docxodus.worker.ts npm/src/worker-proxy.ts | head -40
```

- [ ] **Step 2: Add four routing cases to `docxodus.worker.ts`**

In the worker's message switch, mirror the existing pattern for session methods. For each of the four new methods, add a case that calls the corresponding `session.addAnnotation(...)` (or remove/update/move) and posts back the result. The exact shape mirrors `listAnnotations` — just with the new method names and argument lists.

(If the worker uses a generic `methodName + args` dispatch, no per-method case is needed; just confirm the methods are reachable. Re-run the inspection in Step 1 to verify.)

- [ ] **Step 3: Add proxy methods to `worker-proxy.ts`**

For each new method, add the proxy entry following the existing `addAnnotation`-shaped pattern in that file (or the generic dispatch pattern, same as Step 2).

- [ ] **Step 4: TypeScript compile check**

```bash
cd npm && npx tsc --noEmit
```
Expected: PASS.

- [ ] **Step 5: Build and confirm npm package builds end-to-end**

```bash
cd npm && npm run build
```
Expected: build completes; `dist/` updated.

- [ ] **Step 6: Commit**

```bash
git add npm/src/docxodus.worker.ts npm/src/worker-proxy.ts
git commit -m "feat(npm): route annotation write ops through worker proxy"
```

---

### Task 16: Playwright spec for the WASM wrapper

**Files:**
- Create: `npm/tests/annotations-write.spec.ts`

- [ ] **Step 1: Inspect an existing spec for harness boilerplate**

```bash
ls npm/tests/*.spec.ts | head && grep -n "openDocxSession\|loadFixture" npm/tests/*.spec.ts | head -10
```

- [ ] **Step 2: Create the spec**

```typescript
import { test, expect } from "@playwright/test";
import { openDocxSession } from "../dist/index.js";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

const FIXTURE = resolve(__dirname, "../../TestFiles/DA001-TemplateDocument.docx");

test.describe("annotation write surface", () => {
  test("add → list → remove → save → reopen → empty", async () => {
    const bytes = readFileSync(FIXTURE);
    let session = await openDocxSession(bytes);

    const projection = session.project();
    const firstParagraph = Object.entries(projection.anchorIndex)
      .find(([id, info]) => id.startsWith("p:body:") && (info as any).textPreview);
    expect(firstParagraph).toBeTruthy();
    const [anchorId] = firstParagraph!;

    const addResult = session.addAnnotation(anchorId, { start: 0, length: 1 }, {
      id: "ws-1",
      labelId: "LBL",
      label: "Label",
      color: "#FF0",
      bookmarkName: "", // overwritten by the session
    });
    expect(addResult.success).toBe(true);
    expect(addResult.annotationId).toBe("ws-1");

    expect(session.listAnnotations().some(a => a.id === "ws-1")).toBe(true);

    const rmResult = session.removeAnnotation("ws-1");
    expect(rmResult.success).toBe(true);

    const saved = session.save();
    session.close();

    session = await openDocxSession(saved);
    expect(session.listAnnotations().some(a => a.id === "ws-1")).toBe(false);
    session.close();
  });

  test("update mutates label and metadata", async () => {
    const bytes = readFileSync(FIXTURE);
    const session = await openDocxSession(bytes);
    const projection = session.project();
    const [anchorId] = Object.entries(projection.anchorIndex)
      .find(([id, info]) => id.startsWith("p:body:") && (info as any).textPreview)!;

    session.addAnnotation(anchorId, { start: 0, length: 1 }, {
      id: "ws-2",
      labelId: "OLD", label: "Old", color: "#000", bookmarkName: "",
      metadata: { keep: "yes", drop: "old" },
    });

    const r = session.updateAnnotation("ws-2", {
      label: "New",
      metadataPatch: { drop: null, new: "fresh" },
    });
    expect(r.success).toBe(true);

    const listed = session.listAnnotations().find(a => a.id === "ws-2")!;
    expect(listed.label).toBe("New");
    expect(listed.metadata).toEqual({ keep: "yes", new: "fresh" });
    session.close();
  });

  test("move re-targets the bookmark", async () => {
    const bytes = readFileSync(FIXTURE);
    const session = await openDocxSession(bytes);
    const projection = session.project();
    const paragraphs = Object.entries(projection.anchorIndex)
      .filter(([id, info]) => id.startsWith("p:body:") && (info as any).textPreview)
      .slice(0, 2);
    expect(paragraphs.length).toBe(2);

    session.addAnnotation(paragraphs[0][0], { start: 0, length: 1 }, {
      id: "ws-3", labelId: "L", label: "L", color: "#000", bookmarkName: "",
    });
    const r = session.moveAnnotation("ws-3", paragraphs[1][0], { start: 0, length: 2 });
    expect(r.success).toBe(true);
    expect(r.modified.length).toBe(2);
    session.close();
  });
});
```

- [ ] **Step 3: Build and run the spec**

```bash
cd npm && npm run build && npx playwright test --grep "annotation write surface"
```
Expected: all three tests PASS.

- [ ] **Step 4: Commit**

```bash
git add npm/tests/annotations-write.spec.ts
git commit -m "test(npm): Playwright coverage for annotation write surface"
```

---

## Phase 8: Python (`docx-scalpel`)

### Task 17: Add dispatcher cases

**Files:**
- Modify: `tools/python-host/Dispatcher.cs`

- [ ] **Step 1: Inspect the existing pattern**

```bash
grep -n "list_annotations\|case \"" tools/python-host/Dispatcher.cs | head -20
```

- [ ] **Step 2: Add four switch cases**

Locate the `"list_annotations" => DocxSessionOps.ListAnnotations(Handle(args))` case (line ~94) and add immediately after:

```csharp
        "add_annotation" => DocxSessionOps.AddAnnotation(
            Handle(args),
            Str(args, "anchorId"),
            OptCharSpan(args, "span"),
            Str(args, "annotation")),
        "remove_annotation" => DocxSessionOps.RemoveAnnotation(
            Handle(args), Str(args, "annotationId")),
        "update_annotation" => DocxSessionOps.UpdateAnnotation(
            Handle(args), Str(args, "annotationId"), Str(args, "update")),
        "move_annotation" => DocxSessionOps.MoveAnnotation(
            Handle(args),
            Str(args, "annotationId"),
            Str(args, "newAnchorId"),
            OptCharSpan(args, "newSpan")),
```

If `OptCharSpan` doesn't yet exist as a helper, add it near the existing `Str` / `Handle` helpers:

```csharp
    private static CharSpan? OptCharSpan(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var el) || el.ValueKind == JsonValueKind.Null)
            return null;
        return new CharSpan(
            el.GetProperty("start").GetInt32(),
            el.GetProperty("length").GetInt32());
    }
```

`Str(args, name)` already returns the JSON-stringified subtree when the value is an object — confirm by reading its current implementation:

```bash
grep -n "private static string Str\b" tools/python-host/Dispatcher.cs
```

If `Str` only handles strings (not objects), add a sibling helper `StrOrObject(args, name)` that returns the raw JSON for objects, and use it for the `annotation` and `update` parameters above.

- [ ] **Step 3: Build the host**

```bash
dotnet build tools/python-host/pyhost.csproj
```
Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add tools/python-host/Dispatcher.cs
git commit -m "feat(pyhost): dispatch annotation write ops to DocxSessionOps"
```

---

### Task 18: Extend Python types and add session methods

**Files:**
- Modify: `python/src/docx_scalpel/types.py`
- Modify: `python/src/docx_scalpel/session.py`

- [ ] **Step 1: Extend `DocumentAnnotation` dataclass**

Open `python/src/docx_scalpel/types.py`. Find `class DocumentAnnotation` (line ~421) and add a `metadata` field:

```python
@dataclass(frozen=True, slots=True)
class DocumentAnnotation:
    id: str
    label_id: str
    label: str
    color: str
    bookmark_name: str
    author: str | None = None
    created: str | None = None
    annotated_text: str | None = None
    metadata: Mapping[str, str] = field(default_factory=dict)

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocumentAnnotation":
        return cls(
            id=d["id"],
            label_id=d.get("labelId", ""),
            label=d.get("label", ""),
            color=d.get("color", ""),
            bookmark_name=d.get("bookmarkName", ""),
            author=d.get("author"),
            created=d.get("created"),
            annotated_text=d.get("annotatedText"),
            metadata=dict(d.get("metadata", {}) or {}),
        )

    def to_wire(self) -> dict[str, Any]:
        wire: dict[str, Any] = {
            "id": self.id,
            "labelId": self.label_id,
            "label": self.label,
            "color": self.color,
            "bookmarkName": self.bookmark_name,
        }
        if self.author is not None: wire["author"] = self.author
        if self.created is not None: wire["created"] = self.created
        if self.annotated_text is not None: wire["annotatedText"] = self.annotated_text
        if self.metadata: wire["metadata"] = dict(self.metadata)
        return wire
```

Make sure `field` is imported: `from dataclasses import dataclass, field` at the top.

- [ ] **Step 2: Add `AnnotationUpdate` dataclass**

Add right after `DocumentAnnotation`:

```python
@dataclass(frozen=True, slots=True)
class AnnotationUpdate:
    """Partial-update payload for :meth:`DocxSession.update_annotation`.

    ``None`` / missing fields leave the existing value unchanged.
    ``metadata_patch`` is a per-key merge: a non-``None`` value sets the
    key, an explicit ``None`` removes it, a missing key leaves it
    unchanged.
    """

    label_id: str | None = None
    label: str | None = None
    color: str | None = None
    author: str | None = None
    metadata_patch: Mapping[str, str | None] | None = None

    def to_wire(self) -> dict[str, Any]:
        wire: dict[str, Any] = {}
        if self.label_id is not None: wire["labelId"] = self.label_id
        if self.label is not None: wire["label"] = self.label
        if self.color is not None: wire["color"] = self.color
        if self.author is not None: wire["author"] = self.author
        if self.metadata_patch is not None:
            # Preserve explicit None values — they mean "remove this key".
            wire["metadataPatch"] = dict(self.metadata_patch)
        return wire
```

Re-export both from `python/src/docx_scalpel/__init__.py` (add to the `__all__` if the file defines one, otherwise add them to the public re-export list following the existing pattern).

- [ ] **Step 3: Add four methods on `DocxSession`**

Open `python/src/docx_scalpel/session.py`. After the read-side `list_annotations` method (line ~396), add a new annotations section:

```python
    # -- Tier E: annotations (write surface) -------------------------------

    def add_annotation(
        self,
        anchor_id: str,
        span: CharSpan | None,
        annotation: DocumentAnnotation,
    ) -> EditResult:
        """Annotate a range inside ``anchor_id``.

        When ``span`` is ``None`` the annotation wraps every inline run of
        the block. When ``annotation.id`` is empty, a 16-char hex id is
        auto-generated; check ``EditResult.annotation_id`` for the id used.
        """
        args: dict[str, Any] = {
            "anchorId": anchor_id,
            "annotation": annotation.to_wire(),
        }
        if span is not None:
            args["span"] = {"start": span.start, "length": span.length}
        return EditResult._from_wire(self._call("add_annotation", args))

    def remove_annotation(self, annotation_id: str) -> EditResult:
        return EditResult._from_wire(
            self._call("remove_annotation", {"annotationId": annotation_id})
        )

    def update_annotation(
        self, annotation_id: str, update: AnnotationUpdate,
    ) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "update_annotation",
                {"annotationId": annotation_id, "update": update.to_wire()},
            )
        )

    def move_annotation(
        self, annotation_id: str, new_anchor_id: str, new_span: CharSpan | None,
    ) -> EditResult:
        args: dict[str, Any] = {
            "annotationId": annotation_id,
            "newAnchorId": new_anchor_id,
        }
        if new_span is not None:
            args["newSpan"] = {"start": new_span.start, "length": new_span.length}
        return EditResult._from_wire(self._call("move_annotation", args))
```

Add `AnnotationUpdate` to the imports at the top of `session.py`.

- [ ] **Step 4: Extend `EditResult` to surface `annotation_id`**

Open `python/src/docx_scalpel/types.py`. Find the `EditResult` dataclass (line ~356) and add the new field + parser line:

```python
@dataclass(frozen=True, slots=True)
class EditResult:
    success: bool
    created: tuple[Anchor, ...] = ()
    removed: tuple[Anchor, ...] = ()
    modified: tuple[Anchor, ...] = ()
    patch: MarkdownPatch | None = None
    error: EditError | None = None
    annotation_id: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditResult":
        patch_d = d.get("patch")
        err_d = d.get("error")
        return cls(
            success=bool(d.get("success", False)),
            created=tuple(Anchor._from_wire(a) for a in d.get("created", ())),
            removed=tuple(Anchor._from_wire(a) for a in d.get("removed", ())),
            modified=tuple(Anchor._from_wire(a) for a in d.get("modified", ())),
            patch=MarkdownPatch._from_wire(patch_d) if patch_d else None,
            error=EditError._from_wire(err_d) if err_d else None,
            annotation_id=d.get("annotationId"),
        )
```

- [ ] **Step 5: Type-check Python**

```bash
cd python && .venv/bin/python -m mypy src/docx_scalpel 2>&1 | head -30
```
Expected: no new errors. (If `mypy` isn't configured, skip and rely on tests.)

- [ ] **Step 6: Commit**

```bash
git add python/src/docx_scalpel/types.py python/src/docx_scalpel/session.py python/src/docx_scalpel/__init__.py
git commit -m "feat(docx-scalpel): expose annotation write surface in Python"
```

---

### Task 19: pytest coverage for the Python wrapper

**Files:**
- Create: `python/tests/test_annotations_write.py`

- [ ] **Step 1: Inspect an existing test for fixture loading**

```bash
grep -n "open_session\|fixture\|@pytest" python/tests/test_lifecycle.py | head -20
```

- [ ] **Step 2: Create the spec**

```python
"""End-to-end tests for the annotation write surface through docx-scalpel."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_scalpel import (
    AnnotationUpdate,
    CharSpan,
    DocumentAnnotation,
    open_session,
)

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_paragraph_anchor(session) -> str:
    projection = session.project()
    for anchor_id, info in projection.anchor_index.items():
        if anchor_id.startswith("p:body:") and info.text_preview:
            return anchor_id
    pytest.fail("no body paragraph with text in fixture")


def test_add_and_remove_round_trip():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        r = session.add_annotation(
            anchor,
            CharSpan(start=0, length=1),
            DocumentAnnotation(
                id="py-1", label_id="LBL", label="Lbl",
                color="#FF0", bookmark_name="",
            ),
        )
        assert r.success
        assert r.annotation_id == "py-1"
        assert any(a.id == "py-1" for a in session.list_annotations())

        r = session.remove_annotation("py-1")
        assert r.success
        assert not any(a.id == "py-1" for a in session.list_annotations())


def test_add_with_auto_id():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        r = session.add_annotation(
            anchor, CharSpan(0, 1),
            DocumentAnnotation(id="", label_id="L", label="L", color="#000", bookmark_name=""),
        )
        assert r.success
        assert r.annotation_id is not None
        assert len(r.annotation_id) == 16


def test_add_anchor_not_found_returns_error():
    with open_session(FIXTURE.read_bytes()) as session:
        r = session.add_annotation(
            "p:body:DEADBEEFDEADBEEF", CharSpan(0, 1),
            DocumentAnnotation(id="x", label_id="L", label="L", color="#000", bookmark_name=""),
        )
        assert not r.success
        assert r.error is not None
        assert r.error.code == "AnchorNotFound"


def test_update_mutates_metadata():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        session.add_annotation(
            anchor, CharSpan(0, 1),
            DocumentAnnotation(
                id="py-u", label_id="L", label="L", color="#000", bookmark_name="",
                metadata={"keep": "yes", "drop": "old"},
            ),
        )
        r = session.update_annotation(
            "py-u",
            AnnotationUpdate(label="New", metadata_patch={"drop": None, "new": "fresh"}),
        )
        assert r.success
        listed = next(a for a in session.list_annotations() if a.id == "py-u")
        assert listed.label == "New"
        assert listed.metadata == {"keep": "yes", "new": "fresh"}


def test_move_retargets_modified_includes_both_blocks():
    with open_session(FIXTURE.read_bytes()) as session:
        projection = session.project()
        anchors = [
            a for a, info in projection.anchor_index.items()
            if a.startswith("p:body:") and info.text_preview
        ][:2]
        assert len(anchors) == 2

        session.add_annotation(
            anchors[0], CharSpan(0, 1),
            DocumentAnnotation(id="py-m", label_id="L", label="L", color="#000", bookmark_name=""),
        )
        r = session.move_annotation("py-m", anchors[1], CharSpan(0, 2))
        assert r.success
        modified_ids = {m.id for m in r.modified}
        assert anchors[0] in modified_ids
        assert anchors[1] in modified_ids


def test_save_reopen_persists_annotation():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        session.add_annotation(
            anchor, CharSpan(0, 4),
            DocumentAnnotation(
                id="persist", label_id="P", label="P", color="#0F0", bookmark_name="",
                metadata={"k": "v"},
            ),
        )
        saved = session.save()

    with open_session(saved) as reopened:
        listed = next(a for a in reopened.list_annotations() if a.id == "persist")
        assert listed.label_id == "P"
        assert listed.metadata["k"] == "v"
```

- [ ] **Step 3: Run the tests**

```bash
cd python && .venv/bin/python -m pytest tests/test_annotations_write.py -v
```
Expected: all six tests PASS.

- [ ] **Step 4: Commit**

```bash
git add python/tests/test_annotations_write.py
git commit -m "test(docx-scalpel): pytest coverage for annotation write surface"
```

---

## Phase 9: Documentation

### Task 20: Update CHANGELOG, CLAUDE.md, and architecture docs

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `CLAUDE.md`
- Modify: `docs/architecture/docx_mutation_api.md`
- Modify: `docs/architecture/python_docxodus.md`

- [ ] **Step 1: CHANGELOG entry**

In `CHANGELOG.md`, find the `## [Unreleased]` section and add under `### Added`:

```markdown
- Annotation write surface on `DocxSession` (`AddAnnotation`,
  `RemoveAnnotation`, `UpdateAnnotation`, `MoveAnnotation`) exposed across
  .NET, WASM (`@docxodus/wasm`), and Python (`docx-scalpel`). New
  `EditErrorCode` values: `DuplicateAnnotationId`, `AnnotationNotFound`,
  `EmptyAnnotationSpan`. `EditResult` gained an `AnnotationId` field.
  `AnnotationUpdate` is the new partial-update payload for
  `UpdateAnnotation`.
```

- [ ] **Step 2: CLAUDE.md — update DocxSession bullet**

In `CLAUDE.md`, find the `DocxSession.cs` bullet under "Core Modules" (it lists Tier A/B/C/D). Add Tier E:

```markdown
- Tier E (annotations): `AddAnnotation(anchor, span, annotation)`,
  `RemoveAnnotation(id)`, `UpdateAnnotation(id, AnnotationUpdate)`,
  `MoveAnnotation(id, newAnchor, newSpan)` — anchor-addressed
  annotation CRUD that mutates the live session document.
```

- [ ] **Step 3: CLAUDE.md — update the rippling reminder**

Find the paragraph in CLAUDE.md that begins "When the core library changes a public method or setting on `DocxSession`, update **`Docxodus/Internal/DocxSessionOps.cs` first**" and add `python/src/docx_scalpel/session.py` to the rippling layers list. The existing sentence lists tests / WASM JSExport / stdio dispatcher / `npm/src/types.ts` + `npm/src/index.ts`; insert `python/src/docx_scalpel/session.py + types.py` alongside the npm path.

- [ ] **Step 4: `docs/architecture/docx_mutation_api.md` — add Tier E section**

Append a new section to `docs/architecture/docx_mutation_api.md`:

```markdown
## Tier E: Annotations

Anchor-addressed CRUD for the Docxodus annotation system (custom-XML + bookmark
pairs). Mutates the live session document; round-trips through Save and Reopen.

### Methods

| Method | Description |
|--------|-------------|
| `AddAnnotation(anchorId, span?, DocumentAnnotation)` | Annotate a range; auto-generates 16-char hex id when `annotation.Id` is empty. `BookmarkName`, `AnnotatedText`, `Created`, and `PageInfoStale` are always set by the session. |
| `RemoveAnnotation(id)` | Removes the bookmark pair and the custom-XML entry. |
| `UpdateAnnotation(id, AnnotationUpdate)` | Mutates scalar fields (label/labelId/color/author) and metadata (per-key merge, explicit null = remove key). Range is preserved. |
| `MoveAnnotation(id, newAnchorId, newSpan?)` | Atomically re-targets to a new anchor + span. Validates the new range *before* removing the old bookmark. |

### `AnnotationUpdate`

```csharp
public sealed record AnnotationUpdate
{
    public string? LabelId { get; init; }
    public string? Label { get; init; }
    public string? Color { get; init; }
    public string? Author { get; init; }
    public IReadOnlyDictionary<string, string?>? MetadataPatch { get; init; }
}
```

Null/missing fields leave existing values unchanged. `MetadataPatch` per-key
semantics: non-null value = set/replace, explicit `null` = remove, missing key
= leave as-is.

### New error codes

- `DuplicateAnnotationId` — caller-supplied id already exists, or auto-id
  collided 4 times in a row (vanishingly rare).
- `AnnotationNotFound` — `Remove`/`Update`/`Move` invoked with an unknown id.
- `EmptyAnnotationSpan` — `span.Length == 0`, or `span == null` and the
  resolved block has zero inline runs.

### Return shape

All four ops use the standard `EditResult` envelope with one new field:
`AnnotationId` carries the affected id on success. `Created`/`Removed` are
always empty for these ops (the bookmark + custom-XML entry are internal,
not markdown anchors). `Patch` is null — annotation ops don't change the
markdown projection. `Modified` lists the enclosing block anchor (one entry
for Add/Remove/Update; one or two for Move depending on whether the
destination is the same block as the source).
```

- [ ] **Step 5: `docs/architecture/python_docxodus.md` — add the four ops**

Append the four new methods to whatever method inventory table the file
carries. Keep it brief — one row per method with the Python signature.

- [ ] **Step 6: Build everything end-to-end to confirm nothing regressed**

```bash
dotnet build Docxodus.sln -c Release
dotnet test Docxodus.Tests/Docxodus.Tests.csproj
./scripts/build-wasm.sh
cd npm && npm run build && npx playwright test --grep "annotation write surface"
cd ../python && .venv/bin/python -m pytest tests/test_annotations_write.py -v
```
Expected: all PASS.

- [ ] **Step 7: Commit docs**

```bash
git add CHANGELOG.md CLAUDE.md docs/architecture/docx_mutation_api.md docs/architecture/python_docxodus.md
git commit -m "docs: document annotation write surface (CHANGELOG, CLAUDE.md, architecture)"
```

---

## Done

All 20 tasks complete. The branch `feat/annotation-write-surface` should now contain:
- 2 new internal C# files (`AnnotationsCustomXml.cs`, `AnnotationOps.cs`)
- 1 new xUnit test file with 15+ test methods
- 1 new Playwright spec
- 1 new pytest spec
- Targeted edits across `DocxSession.cs`, `DocxSessionOps.cs`, `DocxSessionJson.cs`, `AnnotationManager.cs`, `DocxSessionBridge.cs`, `Dispatcher.cs`, npm types/wrapper/worker-proxy, Python types/session, CHANGELOG, CLAUDE.md, and two architecture docs.

Open a PR with title `feat: annotation write surface (DocxSession + WASM + docx-scalpel)` and the spec link in the body.
