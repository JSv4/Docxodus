# DOCX→HTML on the Python Surface — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Bring DOCX→HTML conversion to the Python (`docx_scalpel`) surface, both stateless (`convert_docx_to_html`) and session-bound (`DocxSession.to_html`), reusing one shared core renderer that the WASM bridge also delegates to.

**Architecture:** Extract the bytes+options→HTML settings mapping (today inlined only in WASM's `ConvertDocxToHtmlComplete`) into `Docxodus/Internal/HtmlConversionOps.cs`. The WASM shell and the stdio Python host both call it. The host gains two NDJSON ops (`convert_to_html`, `session_to_html`); the Python wrapper gains a module function, a `DocxSession.to_html` method, and an `HtmlOptions` dataclass.

**Tech Stack:** C# / .NET 8 (`Docxodus`, `DocxodusWasm`, `docxodus-pyhost`), Python 3 (`docx_scalpel`), xUnit, pytest.

**Branch:** `feat/python-docx-to-html` (already checked out).

---

## File Structure

- **Create** `Docxodus/Internal/HtmlConversionOps.cs` — `HtmlConversionOptions` DTO + `HtmlConversionOps` static class (stateless + session overloads + base64 image handler). Single owner of the settings mapping.
- **Modify** `wasm/DocxodusWasm/DocumentConverter.cs` — `ConvertDocxToHtmlComplete` builds an `HtmlConversionOptions` and delegates; remove the now-unused private `CreateBase64ImageHandler`.
- **Modify** `tools/python-host/Dispatcher.cs` — two new ops + `ParseHtmlOptions` helper.
- **Modify** `python/src/docx_scalpel/types.py` — `HtmlOptions` dataclass.
- **Modify** `python/src/docx_scalpel/session.py` — `convert_docx_to_html` function + `DocxSession.to_html` method + exports.
- **Create** `Docxodus.Tests/HtmlConversionOpsTests.cs` — .NET unit tests.
- **Create** `python/tests/test_to_html.py` — Python tests.
- **Modify** `CHANGELOG.md`, `docs/architecture/python_docxodus.md`.

---

## Task 1: Core `HtmlConversionOps` (stateless)

**Files:**
- Create: `Docxodus/Internal/HtmlConversionOps.cs`
- Test: `Docxodus.Tests/HtmlConversionOpsTests.cs`

- [ ] **Step 1: Write the failing test**

Create `Docxodus.Tests/HtmlConversionOpsTests.cs`:

```csharp
#nullable enable
using System.IO;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlConversionOpsTests
{
    private static byte[] TourPlanBytes() =>
        File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC001-5DayTourPlanTemplate.docx"));

    [Fact]
    public void HCO001_ConvertBytes_ProducesHtmlWithPrefix()
    {
        var options = new HtmlConversionOptions { CssClassPrefix = "zz-" };

        string html = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), options);

        Assert.Contains("<html", html);
        Assert.Contains("zz-", html);
    }
}
```

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~HCO001"`
Expected: FAIL — compile error, `HtmlConversionOps` / `HtmlConversionOptions` do not exist.

- [ ] **Step 3: Write minimal implementation**

Create `Docxodus/Internal/HtmlConversionOps.cs`:

```csharp
#nullable enable

using System;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Options for <see cref="HtmlConversionOps"/>. Mirrors the parameter set of the
/// WASM <c>DocumentConverter.ConvertDocxToHtmlComplete</c> shell so every surface
/// renders identically. Integer-coded modes match the existing WASM wire contract:
/// CommentRenderMode -1=disabled,0=Endnote,1=Inline,2=Margin;
/// PaginationMode 0=None,1=Paginated; AnnotationLabelMode 0=Above,1=Inline,2=Tooltip,3=None.
/// </summary>
public sealed class HtmlConversionOptions
{
    public string PageTitle { get; init; } = "Document";
    public string CssClassPrefix { get; init; } = "docx-";
    public bool FabricateCssClasses { get; init; } = true;
    public string AdditionalCss { get; init; } = "";
    public int CommentRenderMode { get; init; } = -1;
    public string CommentCssClassPrefix { get; init; } = "comment-";
    public int PaginationMode { get; init; }
    public double PaginationScale { get; init; } = 1.0;
    public string PaginationCssClassPrefix { get; init; } = "page-";
    public bool RenderAnnotations { get; init; }
    public int AnnotationLabelMode { get; init; }
    public string AnnotationCssClassPrefix { get; init; } = "annot-";
    public bool RenderFootnotesAndEndnotes { get; init; }
    public bool RenderHeadersAndFooters { get; init; }
    public bool RenderTrackedChanges { get; init; }
    public bool ShowDeletedContent { get; init; } = true;
    public bool RenderMoveOperations { get; init; } = true;
    public bool RenderUnsupportedContentPlaceholders { get; init; }
    public string? DocumentLanguage { get; init; }
}

/// <summary>
/// Single owner of the DOCX-bytes + <see cref="HtmlConversionOptions"/> →
/// HTML-string mapping. Both the WASM <c>DocumentConverter</c> bridge and the
/// stdio Python host route through here, so render behavior lives in one place.
/// Throws on invalid input; callers serialize errors at their boundary.
/// </summary>
public static class HtmlConversionOps
{
    /// <summary>Render raw DOCX bytes to a self-contained HTML string.</summary>
    public static string ConvertToHtml(byte[] docxBytes, HtmlConversionOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(docxBytes));

        // Writable stream required: WmlToHtmlConverter runs RevisionAccepter internally.
        using var memoryStream = new MemoryStream();
        memoryStream.Write(docxBytes, 0, docxBytes.Length);
        memoryStream.Position = 0;
        using var wordDoc = WordprocessingDocument.Open(memoryStream, true);

        var renderComments = options.CommentRenderMode >= 0;

        var settings = new WmlToHtmlConverterSettings
        {
            PageTitle = options.PageTitle,
            CssClassPrefix = options.CssClassPrefix,
            FabricateCssClasses = options.FabricateCssClasses,
            AdditionalCss = options.AdditionalCss,
            GeneralCss = "body { font-family: Arial, sans-serif; margin: 20px; } " +
                         "span { white-space: pre-wrap; }",
            RenderComments = renderComments,
            CommentRenderMode = renderComments
                ? (CommentRenderMode)options.CommentRenderMode
                : CommentRenderMode.EndnoteStyle,
            CommentCssClassPrefix = options.CommentCssClassPrefix,
            IncludeCommentMetadata = true,
            RenderPagination = (PaginationMode)options.PaginationMode,
            PaginationScale = options.PaginationScale > 0 ? options.PaginationScale : 1.0,
            PaginationCssClassPrefix = options.PaginationCssClassPrefix,
            RenderAnnotations = options.RenderAnnotations,
            AnnotationLabelMode = (AnnotationLabelMode)options.AnnotationLabelMode,
            AnnotationCssClassPrefix = options.AnnotationCssClassPrefix,
            IncludeAnnotationMetadata = true,
            RenderFootnotesAndEndnotes = options.RenderFootnotesAndEndnotes,
            RenderHeadersAndFooters = options.RenderHeadersAndFooters,
            RenderTrackedChanges = options.RenderTrackedChanges,
            ShowDeletedContent = options.ShowDeletedContent,
            RenderMoveOperations = options.RenderMoveOperations,
            IncludeRevisionMetadata = true,
            RenderUnsupportedContentPlaceholders = options.RenderUnsupportedContentPlaceholders,
            UnsupportedContentCssClassPrefix = "unsupported-",
            IncludeUnsupportedContentMetadata = true,
            DocumentLanguage = options.DocumentLanguage,
            // Embed images as base64 data URIs — no SkiaSharp needed (WASM-safe).
            ImageHandler = CreateBase64ImageHandler(),
        };

        var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
        return htmlElement.ToString(SaveOptions.DisableFormatting);
    }

    private static Func<ImageInfo, XElement> CreateBase64ImageHandler()
    {
        return imageInfo =>
        {
            if (imageInfo.ImageBytes == null || imageInfo.ImageBytes.Length == 0)
                return null!;

            var mimeType = imageInfo.ContentType ?? "image/png";
            var base64 = Convert.ToBase64String(imageInfo.ImageBytes);
            var dataUri = $"data:{mimeType};base64,{base64}";

            var imgElement = new XElement(XhtmlNoNamespace.img,
                new XAttribute("src", dataUri));

            if (imageInfo.ImgStyleAttribute != null)
                imgElement.Add(imageInfo.ImgStyleAttribute);

            if (!string.IsNullOrEmpty(imageInfo.AltText))
                imgElement.Add(new XAttribute("alt", imageInfo.AltText));

            return imgElement;
        };
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~HCO001"`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add Docxodus/Internal/HtmlConversionOps.cs Docxodus.Tests/HtmlConversionOpsTests.cs
git commit -m "feat(core): HtmlConversionOps shared DOCX->HTML renderer"
```

---

## Task 2: Core `HtmlConversionOps` session overloads

**Files:**
- Modify: `Docxodus/Internal/HtmlConversionOps.cs`
- Test: `Docxodus.Tests/HtmlConversionOpsTests.cs`

- [ ] **Step 1: Write the failing test**

Append to `HtmlConversionOpsTests.cs` (inside the class):

```csharp
    [Fact]
    public void HCO002_ConvertSession_ReflectsEdit()
    {
        using var session = new DocxSession(TourPlanBytes());
        var projection = session.Project();

        // First body paragraph/heading/list-item anchor, in document order.
        // C# AnchorTarget nests the anchor: record struct Anchor(Id, Kind, Scope, Unid).
        string FirstAnchor()
        {
            string? best = null;
            int bestPos = int.MaxValue;
            foreach (var target in projection.AnchorIndex.Values)
            {
                if (target.Anchor.Scope != "body") continue;
                if (target.Anchor.Kind is not ("p" or "h" or "li")) continue;
                int pos = projection.Markdown.IndexOf("{#" + target.Anchor.Id + "}", System.StringComparison.Ordinal);
                if (pos >= 0 && pos < bestPos) { bestPos = pos; best = target.Anchor.Id; }
            }
            Assert.NotNull(best);
            return best!;
        }

        var edit = session.ReplaceText(FirstAnchor(), "HCO002UNIQUEMARKER edited body.");
        Assert.True(edit.Success, edit.Error?.Message);

        string html = HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions());

        Assert.Contains("HCO002UNIQUEMARKER", html);
    }
```

Verified names: `MarkdownProjection.Markdown` / `.AnchorIndex` (dict keyed by anchor id → `AnchorTarget`); `AnchorTarget.Anchor` is `record struct Anchor(string Id, string Kind, string Scope, string Unid)`; `EditResult.Success` / `.Error`; `EditError` is `record EditError(EditErrorCode Code, string Message, string? AnchorId)`.

- [ ] **Step 2: Run test to verify it fails**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~HCO002"`
Expected: FAIL — no `ConvertToHtml(DocxSession, ...)` overload.

- [ ] **Step 3: Write minimal implementation**

In `HtmlConversionOps.cs`, add these two overloads to the `HtmlConversionOps` class (after the `byte[]` overload):

```csharp
    /// <summary>Render a live session's current (possibly edited) state to HTML.</summary>
    public static string ConvertToHtml(DocxSession session, HtmlConversionOptions options)
    {
        if (session == null) throw new ArgumentNullException(nameof(session));
        return ConvertToHtml(session.Save(), options);
    }

    /// <summary>Render the session registered under <paramref name="handle"/> to HTML.</summary>
    public static string ConvertToHtml(int handle, HtmlConversionOptions options) =>
        ConvertToHtml(SessionRegistry.Get(handle), options);
```

- [ ] **Step 4: Run test to verify it passes**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~HtmlConversionOpsTests"`
Expected: PASS (both HCO001 and HCO002).

- [ ] **Step 5: Commit**

```bash
git add Docxodus/Internal/HtmlConversionOps.cs Docxodus.Tests/HtmlConversionOpsTests.cs
git commit -m "feat(core): HtmlConversionOps session + handle overloads"
```

---

## Task 3: WASM dedupe — delegate `ConvertDocxToHtmlComplete`

**Files:**
- Modify: `wasm/DocxodusWasm/DocumentConverter.cs` (the body of `ConvertDocxToHtmlComplete`, ~line 229; and remove the private `CreateBase64ImageHandler`, ~line 1996)

- [ ] **Step 1: Replace the method body**

In `ConvertDocxToHtmlComplete`, replace the `try { ... }` block's settings-construction + conversion (everything from `using var memoryStream` through `return htmlElement.ToString(...)`) with a delegation. The final method body becomes:

```csharp
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            var options = new Docxodus.Internal.HtmlConversionOptions
            {
                PageTitle = pageTitle ?? "Document",
                CssClassPrefix = cssPrefix ?? "docx-",
                FabricateCssClasses = fabricateClasses,
                AdditionalCss = additionalCss ?? "",
                CommentRenderMode = commentRenderMode,
                CommentCssClassPrefix = commentCssClassPrefix ?? "comment-",
                PaginationMode = paginationMode,
                PaginationScale = paginationScale,
                PaginationCssClassPrefix = paginationCssClassPrefix ?? "page-",
                RenderAnnotations = renderAnnotations,
                AnnotationLabelMode = annotationLabelMode,
                AnnotationCssClassPrefix = annotationCssClassPrefix ?? "annot-",
                RenderFootnotesAndEndnotes = renderFootnotesAndEndnotes,
                RenderHeadersAndFooters = renderHeadersAndFooters,
                RenderTrackedChanges = renderTrackedChanges,
                ShowDeletedContent = showDeletedContent,
                RenderMoveOperations = renderMoveOperations,
                RenderUnsupportedContentPlaceholders = renderUnsupportedContentPlaceholders,
                DocumentLanguage = documentLanguage,
            };
            return Docxodus.Internal.HtmlConversionOps.ConvertToHtml(docxBytes, options);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
```

> Note: this changes the default `CssClassPrefix` fallback to `"docx-"` (was already `"docx-"` in the original) — unchanged. The old body's `PaginationScale = paginationScale > 0 ? paginationScale : 1.0` guard now lives inside `HtmlConversionOps`, so passing `paginationScale` straight through is correct.

- [ ] **Step 2: Remove the now-unused private helper**

Delete the entire `private static Func<ImageInfo, XElement> CreateBase64ImageHandler()` method (~line 1996) from `DocumentConverter.cs` **only if** no other method still references it.

Run: `grep -n "CreateBase64ImageHandler" wasm/DocxodusWasm/DocumentConverter.cs`
- If matches remain (other converters like the annotation/external-annotation paths use it), **leave it in place** and skip this step.
- If the only match was the definition, delete the method.

- [ ] **Step 3: Build to verify (non-WASM compile of the shared lib + WASM project compiles)**

Run: `dotnet build Docxodus/Docxodus.csproj`
Expected: Build succeeded.

Then the WASM build:
Run: `./scripts/build-wasm.sh`
Expected: completes without error.

> If you switch back to `dotnet test` after a WASM build, run `dotnet clean Docxodus.sln` first (see CLAUDE.md "Switching back from a WASM build").

- [ ] **Step 4: WASM regression test**

Run:
```bash
cd npm && npm run build && npm test -- --grep "convert" 2>/dev/null || npm test
```
Expected: existing DOCX→HTML conversion specs PASS (behavior unchanged by the dedupe).

- [ ] **Step 5: Commit**

```bash
cd /home/jman/Code/Docxodus
git add wasm/DocxodusWasm/DocumentConverter.cs
git commit -m "refactor(wasm): ConvertDocxToHtmlComplete delegates to HtmlConversionOps"
```

---

## Task 4: Python host ops — `convert_to_html`, `session_to_html`

**Files:**
- Modify: `tools/python-host/Dispatcher.cs`

- [ ] **Step 1: Add the two op cases**

In the `Dispatch` switch (after the `"save"` case near line 30), add:

```csharp
        "convert_to_html" => ConvertToHtml(args),
        "session_to_html" => SessionToHtml(args),
```

- [ ] **Step 2: Add the op + options-parsing helpers**

Add these private methods to `Dispatcher` (place near `Save`):

```csharp
    private static string ConvertToHtml(JsonElement args)
    {
        var bytes = Convert.FromBase64String(Str(args, "docxB64"));
        var html = HtmlConversionOps.ConvertToHtml(bytes, ParseHtmlOptions(args));
        return JsonString(html);
    }

    private static string SessionToHtml(JsonElement args)
    {
        var html = HtmlConversionOps.ConvertToHtml(Handle(args), ParseHtmlOptions(args));
        return JsonString(html);
    }

    private static HtmlConversionOptions ParseHtmlOptions(JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object
            || !args.TryGetProperty("options", out var o)
            || o.ValueKind != JsonValueKind.Object)
        {
            return new HtmlConversionOptions();
        }

        string StrOpt(string name, string fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String
                ? v.GetString()! : fallback;
        int IntOpt(string name, int fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
                ? v.GetInt32() : fallback;
        double DblOpt(string name, double fallback) =>
            o.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number
                ? v.GetDouble() : fallback;
        bool BoolOpt(string name, bool fallback) =>
            o.TryGetProperty(name, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False)
                ? v.GetBoolean() : fallback;

        var defaults = new HtmlConversionOptions();
        return new HtmlConversionOptions
        {
            PageTitle = StrOpt("pageTitle", defaults.PageTitle),
            CssClassPrefix = StrOpt("cssClassPrefix", defaults.CssClassPrefix),
            FabricateCssClasses = BoolOpt("fabricateCssClasses", defaults.FabricateCssClasses),
            AdditionalCss = StrOpt("additionalCss", defaults.AdditionalCss),
            CommentRenderMode = IntOpt("commentRenderMode", defaults.CommentRenderMode),
            CommentCssClassPrefix = StrOpt("commentCssClassPrefix", defaults.CommentCssClassPrefix),
            PaginationMode = IntOpt("paginationMode", defaults.PaginationMode),
            PaginationScale = DblOpt("paginationScale", defaults.PaginationScale),
            PaginationCssClassPrefix = StrOpt("paginationCssClassPrefix", defaults.PaginationCssClassPrefix),
            RenderAnnotations = BoolOpt("renderAnnotations", defaults.RenderAnnotations),
            AnnotationLabelMode = IntOpt("annotationLabelMode", defaults.AnnotationLabelMode),
            AnnotationCssClassPrefix = StrOpt("annotationCssClassPrefix", defaults.AnnotationCssClassPrefix),
            RenderFootnotesAndEndnotes = BoolOpt("renderFootnotesAndEndnotes", defaults.RenderFootnotesAndEndnotes),
            RenderHeadersAndFooters = BoolOpt("renderHeadersAndFooters", defaults.RenderHeadersAndFooters),
            RenderTrackedChanges = BoolOpt("renderTrackedChanges", defaults.RenderTrackedChanges),
            ShowDeletedContent = BoolOpt("showDeletedContent", defaults.ShowDeletedContent),
            RenderMoveOperations = BoolOpt("renderMoveOperations", defaults.RenderMoveOperations),
            RenderUnsupportedContentPlaceholders = BoolOpt("renderUnsupportedContentPlaceholders", defaults.RenderUnsupportedContentPlaceholders),
            DocumentLanguage = o.TryGetProperty("documentLanguage", out var dl) && dl.ValueKind == JsonValueKind.String
                ? dl.GetString() : defaults.DocumentLanguage,
        };
    }
```

`HtmlConversionOptions` / `HtmlConversionOps` resolve via the existing `using Docxodus.Internal;` at the top of the file.

- [ ] **Step 3: Build the host**

Run: `dotnet build tools/python-host/docxodus-pyhost.csproj`
Expected: Build succeeded. (If the csproj name differs, run `ls tools/python-host/*.csproj` and use that path.)

- [ ] **Step 4: Commit**

```bash
git add tools/python-host/Dispatcher.cs
git commit -m "feat(pyhost): convert_to_html + session_to_html ops"
```

---

## Task 5: Python wrapper — `HtmlOptions`, `convert_docx_to_html`, `to_html`

**Files:**
- Modify: `python/src/docx_scalpel/types.py`
- Modify: `python/src/docx_scalpel/session.py`

- [ ] **Step 1: Add the `HtmlOptions` dataclass**

In `python/src/docx_scalpel/types.py`, add (near the other frozen dataclasses; ensure `from dataclasses import dataclass, field` is imported — check the file's existing imports):

```python
@dataclass(frozen=True, slots=True)
class HtmlOptions:
    """Options for DOCX→HTML conversion. Mirrors the .NET ``HtmlConversionOptions``.

    Integer-coded modes match the wire contract:
    ``comment_render_mode`` -1=disabled,0=endnote,1=inline,2=margin;
    ``pagination_mode`` 0=none,1=paginated;
    ``annotation_label_mode`` 0=above,1=inline,2=tooltip,3=none.
    """

    page_title: str = "Document"
    css_class_prefix: str = "docx-"
    fabricate_css_classes: bool = True
    additional_css: str = ""
    comment_render_mode: int = -1
    comment_css_class_prefix: str = "comment-"
    pagination_mode: int = 0
    pagination_scale: float = 1.0
    pagination_css_class_prefix: str = "page-"
    render_annotations: bool = False
    annotation_label_mode: int = 0
    annotation_css_class_prefix: str = "annot-"
    render_footnotes_and_endnotes: bool = False
    render_headers_and_footers: bool = False
    render_tracked_changes: bool = False
    show_deleted_content: bool = True
    render_move_operations: bool = True
    render_unsupported_content_placeholders: bool = False
    document_language: str | None = None

    def to_wire(self) -> dict[str, Any]:
        """camelCase keys the host dispatcher's ``ParseHtmlOptions`` reads."""
        wire: dict[str, Any] = {
            "pageTitle": self.page_title,
            "cssClassPrefix": self.css_class_prefix,
            "fabricateCssClasses": self.fabricate_css_classes,
            "additionalCss": self.additional_css,
            "commentRenderMode": self.comment_render_mode,
            "commentCssClassPrefix": self.comment_css_class_prefix,
            "paginationMode": self.pagination_mode,
            "paginationScale": self.pagination_scale,
            "paginationCssClassPrefix": self.pagination_css_class_prefix,
            "renderAnnotations": self.render_annotations,
            "annotationLabelMode": self.annotation_label_mode,
            "annotationCssClassPrefix": self.annotation_css_class_prefix,
            "renderFootnotesAndEndnotes": self.render_footnotes_and_endnotes,
            "renderHeadersAndFooters": self.render_headers_and_footers,
            "renderTrackedChanges": self.render_tracked_changes,
            "showDeletedContent": self.show_deleted_content,
            "renderMoveOperations": self.render_move_operations,
            "renderUnsupportedContentPlaceholders": self.render_unsupported_content_placeholders,
        }
        if self.document_language is not None:
            wire["documentLanguage"] = self.document_language
        return wire
```

If `types.py` does not already import `Any`, add `from typing import Any` (verify existing imports first). Add `"HtmlOptions"` to the module's `__all__` if it has one.

- [ ] **Step 2: Wire it through `session.py`**

In `python/src/docx_scalpel/session.py`:

(a) Add `HtmlOptions` to the `from .types import (...)` block.

(b) Update `__all__`:

```python
__all__ = ["DocxSession", "open_session", "ping", "convert_docx_to_html"]
```

(c) Add the module-level function (place near `open_session`):

```python
def convert_docx_to_html(
    data: bytes, options: HtmlOptions | None = None
) -> str:
    """Convert DOCX bytes to a self-contained HTML string (stateless).

    Mirrors the WASM/npm ``convertDocxToHtml``. Opens a throwaway session in the
    host, renders, and discards it — use :meth:`DocxSession.to_html` to render the
    current state of an already-open editing session instead.
    """
    args: dict[str, Any] = {"docxB64": base64.b64encode(data).decode("ascii")}
    if options is not None:
        args["options"] = options.to_wire()
    html = _call("convert_to_html", args)
    if not isinstance(html, str):
        raise TypeError(f"convert_to_html: expected str, got {type(html).__name__}")
    return html
```

(d) Add the session method (inside `class DocxSession`, near `save`):

```python
    def to_html(self, options: HtmlOptions | None = None) -> str:
        """Render this session's current (possibly edited) state to HTML."""
        args: dict[str, Any] = {"handle": self._handle}
        if options is not None:
            args["options"] = options.to_wire()
        html = _call("session_to_html", args)
        if not isinstance(html, str):
            raise TypeError(f"session_to_html: expected str, got {type(html).__name__}")
        return html
```

- [ ] **Step 3: Re-export from the package root if applicable**

Run: `grep -n "convert_docx_to_html\|open_session\|HtmlOptions\|__all__" python/src/docx_scalpel/__init__.py`
If `__init__.py` re-exports `open_session` / public types, add `convert_docx_to_html` and `HtmlOptions` to its imports and `__all__` following the existing pattern.

- [ ] **Step 4: Type-check / import smoke**

Run: `cd python && python -c "from docx_scalpel import convert_docx_to_html, HtmlOptions; print('ok')"`
Expected: `ok` (no ImportError).

- [ ] **Step 5: Commit**

```bash
cd /home/jman/Code/Docxodus
git add python/src/docx_scalpel/types.py python/src/docx_scalpel/session.py python/src/docx_scalpel/__init__.py
git commit -m "feat(python): convert_docx_to_html + DocxSession.to_html"
```

---

## Task 6: Python tests

**Files:**
- Create: `python/tests/test_to_html.py`

- [ ] **Step 1: Write the tests**

Create `python/tests/test_to_html.py`:

```python
"""DOCX→HTML conversion on the Python surface — stateless + session-bound."""

from __future__ import annotations

from docx_scalpel import HtmlOptions, convert_docx_to_html, open_session


def test_th001_stateless_convert_produces_html(tour_plan_bytes: bytes) -> None:
    html = convert_docx_to_html(tour_plan_bytes)
    assert "<html" in html
    assert "</html>" in html


def test_th002_css_prefix_option_applied(tour_plan_bytes: bytes) -> None:
    html = convert_docx_to_html(tour_plan_bytes, HtmlOptions(css_class_prefix="zz-"))
    assert "zz-" in html


def test_th003_session_to_html_reflects_edit(tour_plan_bytes: bytes) -> None:
    marker = "TH003UNIQUEMARKER"
    with open_session(tour_plan_bytes) as session:
        projection = session.project()
        # First body paragraph/heading/list-item anchor in document order.
        candidates = [
            t
            for t in projection.anchor_index.values()
            if t.kind in ("p", "h", "li") and t.scope == "body"
        ]
        anchor = min(
            candidates,
            key=lambda t: (
                projection.markdown.find("{#" + t.id + "}")
                if projection.markdown.find("{#" + t.id + "}") >= 0
                else 1 << 30
            ),
        )
        result = session.replace_text(anchor.id, f"{marker} edited body.")
        assert result.success, result.error

        edited_html = session.to_html()
        assert marker in edited_html

    # Stateless conversion of the ORIGINAL bytes must not contain the edit.
    original_html = convert_docx_to_html(tour_plan_bytes)
    assert marker not in original_html
```

- [ ] **Step 2: Run the tests**

Run: `cd python && python -m pytest tests/test_to_html.py -v`
Expected: 3 passed. (Tests build/launch the host; if the host binary is stale, the suite's conftest/transport rebuilds or points at it — confirm `dotnet build tools/python-host/*.csproj` from Task 4 ran.)

- [ ] **Step 3: Run the full Python suite (regression)**

Run: `cd python && python -m pytest -q`
Expected: all pass.

- [ ] **Step 4: Commit**

```bash
cd /home/jman/Code/Docxodus
git add python/tests/test_to_html.py
git commit -m "test(python): DOCX->HTML stateless + session_to_html"
```

---

## Task 7: Documentation

**Files:**
- Modify: `CHANGELOG.md`
- Modify: `docs/architecture/python_docxodus.md`

- [ ] **Step 1: CHANGELOG entry**

Under the `[Unreleased]` section's `### Added` (create the heading if absent), add:

```markdown
- **Python DOCX→HTML conversion** — `convert_docx_to_html(data, options)` and
  `DocxSession.to_html(options)` in `docx_scalpel`, backed by a new shared
  `HtmlConversionOps` core renderer that the WASM bridge now also delegates to.
  New `HtmlOptions` dataclass mirrors the existing WASM/npm conversion options.
  New stdio-host ops: `convert_to_html`, `session_to_html`.
```

- [ ] **Step 2: Architecture doc**

In `docs/architecture/python_docxodus.md`, add a section documenting the two ops, the `HtmlOptions` fields + defaults, and that `to_html` renders live edited state while `convert_docx_to_html` is stateless. Note that `HtmlConversionOps` is the single shared renderer for both WASM and the host.

- [ ] **Step 3: Verify CLAUDE.md op list (if present)**

Run: `grep -n "convert_to_html\|session_to_html\|to_html" CLAUDE.md`
The `DocxSession.cs` bullet in CLAUDE.md lists inspection ops; if it enumerates host ops, add the two new ops. Otherwise no change needed.

- [ ] **Step 4: Commit**

```bash
git add CHANGELOG.md docs/architecture/python_docxodus.md CLAUDE.md
git commit -m "docs: Python DOCX->HTML conversion surface"
```

---

## Final Verification

- [ ] `dotnet clean Docxodus.sln` (if a WASM build ran since the last `dotnet test`)
- [ ] `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~HtmlConversionOpsTests"` → PASS
- [ ] `cd python && python -m pytest -q` → all PASS
- [ ] `dotnet build -c Release Docxodus.sln` → no warnings-as-errors regressions
- [ ] Review the full diff: `git diff main...feat/python-docx-to-html --stat`
