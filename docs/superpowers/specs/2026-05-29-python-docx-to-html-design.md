# Design: DOCX→HTML conversion on the Python surface

**Date:** 2026-05-29
**Status:** Approved (pending spec review)
**Branch:** `feat/python-docx-to-html`

## Background

This work is the first step toward a "convert a DOCX to images" capability. The
agreed image path is DOCX → HTML → screenshot, with the screenshot engine kept
out of the core library. Rather than build the whole pipeline at once, we land
the HTML foundation first.

DOCX→HTML already ships on three of the four surfaces:

| Surface | Existing entry point |
|---------|----------------------|
| .NET core | `WmlToHtmlConverter.ConvertToHtml(WordprocessingDocument, WmlToHtmlConverterSettings)` + the `docx2html` CLI tool |
| WASM | `DocumentConverter.ConvertDocxToHtml` / `WithOptions` / `WithPagination` / `Full` / `Complete` |
| npm/TS | `convertDocxToHtml(file, options)` |
| **Python (`docx_scalpel`)** | **— none —** |

The Python package is entirely session-handle-based (`DocxSessionOps`,
`Dispatcher`, `DocxSession`) and has no DOCX→HTML path. This design fills that
gap with parity options, and in doing so centralizes the bytes+options→HTML
settings mapping that currently lives only in the WASM shell.

## Goals

1. Python callers can convert DOCX bytes to an HTML string with the same option
   set the other surfaces expose.
2. Python callers can render the **current (possibly edited)** state of an open
   `DocxSession` to HTML.
3. The settings-mapping logic lives in exactly one place in the core library,
   shared by both the WASM bridge and the stdio host — consistent with the
   `DocxSessionOps` / `DocxSessionJson` "wire shapes in one place" principle.

## Non-goals (deferred)

- The screenshot / page-rasterization step (the eventual "to images" feature).
- Any new WASM `[JSExport]` method or npm function — those surfaces are already
  complete and their public behavior does not change.
- A Python CLI binary. `docx_scalpel` is a library; the existing `docx2html`
  .NET CLI covers the command-line use case.

## Architecture

### 1. Core — shared conversion helper

New file `Docxodus/Internal/HtmlConversionOps.cs` (sits beside
`DocxSessionOps.cs`). It owns the canonical bytes+options→HTML mapping that is
currently inlined in `wasm/DocxodusWasm/DocumentConverter.cs::ConvertDocxToHtmlComplete`.

Shape:

```csharp
#nullable enable
namespace Docxodus.Internal;

public sealed class HtmlConversionOptions
{
    public string PageTitle { get; init; } = "Document";
    public string CssClassPrefix { get; init; } = "docx-";
    public bool FabricateCssClasses { get; init; } = true;
    public string AdditionalCss { get; init; } = "";
    public int CommentRenderMode { get; init; } = -1;   // -1=disabled,0=Endnote,1=Inline,2=Margin
    public string CommentCssClassPrefix { get; init; } = "comment-";
    public int PaginationMode { get; init; } = 0;        // 0=None,1=Paginated
    public double PaginationScale { get; init; } = 1.0;
    public string PaginationCssClassPrefix { get; init; } = "page-";
    public bool RenderAnnotations { get; init; }
    public int AnnotationLabelMode { get; init; }        // 0=Above,1=Inline,2=Tooltip,3=None
    public string AnnotationCssClassPrefix { get; init; } = "annot-";
    public bool RenderFootnotesAndEndnotes { get; init; }
    public bool RenderHeadersAndFooters { get; init; }
    public bool RenderTrackedChanges { get; init; }
    public bool ShowDeletedContent { get; init; } = true;
    public bool RenderMoveOperations { get; init; } = true;
    public bool RenderUnsupportedContentPlaceholders { get; init; }
    public string? DocumentLanguage { get; init; }
}

public static class HtmlConversionOps
{
    // Stateless: parse bytes, render. Used by WASM and the host's convert_to_html.
    public static string ConvertToHtml(byte[] docxBytes, HtmlConversionOptions options);

    // Session-bound: render the live session's current Save() bytes.
    // Used by the host's session_to_html.
    public static string ConvertToHtml(int sessionHandle, HtmlConversionOptions options);
}
```

The single private core builds the `WmlToHtmlConverterSettings` exactly as
`ConvertDocxToHtmlComplete` does today — including the base64 image handler
(`CreateBase64ImageHandler()`), so output is self-contained and the WASM path
needs no SkiaSharp. The session overload calls `SessionRegistry.Get(handle)`,
takes its current `Save()` bytes, and routes into the same stateless core.

The base64 image-handler helper currently lives as a private method in
`DocumentConverter.cs`. It moves into `HtmlConversionOps` as the single owner;
`DocumentConverter` no longer needs its own copy once it delegates. Errors
propagate as thrown exceptions — the bridges keep their existing JSON-error
serialization at the boundary.

### 2. WASM refactor (dedupe)

`DocumentConverter.ConvertDocxToHtmlComplete` is rewritten to build an
`HtmlConversionOptions` from its parameters and delegate to
`HtmlConversionOps.ConvertToHtml(bytes, options)`. The five public
`[JSExport]` signatures and their JSON-error behavior are unchanged — this is a
pure internal dedupe so WASM and Python share one renderer. Existing WASM tests
guard the behavior.

### 3. Python host — `tools/python-host/Dispatcher.cs`

Two new NDJSON ops:

| Op | Args | Routes to |
|----|------|-----------|
| `convert_to_html` | `docxB64` (base64 string), `options` (object, optional) | `HtmlConversionOps.ConvertToHtml(bytes, options)` |
| `session_to_html` | `handle` (int), `options` (object, optional) | `HtmlConversionOps.ConvertToHtml(handle, options)` |

A small `ParseHtmlOptions(JsonElement)` helper in the dispatcher decodes the
`options` object into `HtmlConversionOptions`, applying the same defaults as the
DTO when a field is absent. Both ops return the raw HTML string via the existing
`JsonString(...)` wrapper (the wire carries it as a JSON string value).

### 4. Python wrapper — `docx_scalpel`

`types.py`: new `@dataclass(frozen=True, slots=True) HtmlOptions` mirroring the
DTO fields (snake_case), with matching defaults and a `to_wire()` that emits the
camelCase keys the dispatcher reads. Enums (`CommentRenderMode`,
`PaginationMode`, `AnnotationLabelMode`) may be modeled as `IntEnum`s in
`enums.py` for ergonomics, but the wire values stay integers.

`session.py`:
- Module-level `convert_docx_to_html(data: bytes, options: HtmlOptions | None = None) -> str`
  — base64-encodes `data`, calls `_call("convert_to_html", {...})`. Added to
  `__all__`.
- `DocxSession.to_html(self, options: HtmlOptions | None = None) -> str`
  — calls `_call("session_to_html", {"handle": self._handle, ...})`; renders
  current edited state.

## Data flow

```
convert_docx_to_html(bytes, opts)
  → NDJSON {op:"convert_to_html", docxB64, options}
  → Dispatcher.ParseHtmlOptions → HtmlConversionOps.ConvertToHtml(bytes, opts)
  → WmlToHtmlConverter.ConvertToHtml → html string → JSON string → Python str

session.to_html(opts)
  → NDJSON {op:"session_to_html", handle, options}
  → SessionRegistry.Get(handle).Save() → HtmlConversionOps.ConvertToHtml(bytes, opts)
  → html string → Python str
```

## Error handling

- Empty/null bytes → host throws; dispatcher surfaces the existing error
  envelope; Python `_transport` raises as it does for every other op.
- Unknown session handle on `session_to_html` → `SessionRegistry.Get` throws the
  same "unknown handle" error used by all session ops.
- Malformed `options` object → dispatcher falls back to per-field defaults
  rather than failing (absent field = default), matching how the other surfaces
  treat omitted parameters.

## Testing

- **.NET** (`Docxodus.Tests`): a test for `HtmlConversionOps.ConvertToHtml(bytes, options)`
  asserting key structural markers (e.g. `<html`, the css prefix, a known
  paragraph's text) on an existing `TestFiles/HC*` document. A second test that
  opens a session, mutates a block, and asserts the session overload's HTML
  reflects the edit.
- **WASM**: existing `DocumentConverter` HTML tests must still pass unchanged
  (regression guard for the refactor) — run `npm run build` then `npm test`.
- **Python** (`python/tests`): round-trip a known DOCX through
  `convert_docx_to_html` and assert structural markers; open a session, edit a
  paragraph, and assert `session.to_html()` contains the new text and the
  stateless conversion of the original bytes does not.

## Documentation

- `CHANGELOG.md`: entry under `[Unreleased]`.
- `docs/architecture/python_docxodus.md`: document the two new ops, `HtmlOptions`,
  and the `to_html` / `convert_docx_to_html` surface.
- `CLAUDE.md`: add the new ops to the python-host dispatcher description if the
  op list is enumerated there (verify during implementation).

## Open question resolved during design

- **Python shape:** both a stateless `convert_docx_to_html` and a
  `session.to_html` (chosen over either alone) — parity with other surfaces plus
  the ability to render mid-edit state.
