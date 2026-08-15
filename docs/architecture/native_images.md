# Native image session API

Issue #453 adds an occurrence-oriented image API to `DocxSession`. It edits OOXML picture
markup and package relationships directly; it is not an image decoder, URL downloader, or
filesystem facade. Call `DocxSession.GetImageCapabilities()` (or the equivalent JSON/client
method) when behavior must be selected at runtime.

## Public contract

`ListImages(scopes)` walks the body, every header/footer, footnotes, endnotes, and comments through the
shared owning-part seam. It returns one `ImageOccurrence` for every DrawingML `a:blip` and legacy
VML `v:imagedata` it can identify. An occurrence reports:

- a stable story-scoped id, anchor, and zero-length character span at the picture boundary;
- the owning part and its owner-local embedded or external relationship id/target;
- markup/placement kind, intrinsic pixel dimensions, rendered point dimensions, media filename,
  declared content type, signature-detected format, and content-type/signature agreement;
- alt text/title and typed floating layout facts; and
- `canMutate` plus `unsupportedReason`, so inspection never implies write support.

One non-canonical drawing can contain several blips. Each is returned separately with `:subN`
on the common drawing id. These rows are read-only. A drawing with no identifiable blip is also
reported as unsupported instead of disappearing — that guarantee is scoped to drawings the
projection can address, i.e. those inside a `w:p` that resolves to an anchor. A `w:drawing`
whose nearest paragraph has no anchor is skipped entirely and appears in no listing. IDs are
structural occurrence IDs, not media part IDs: two occurrences may deliberately share one
media part.

The mutation surface is `InsertImage`, `ReplaceImage`, `SetImageDimensions`,
`SetImageMetadata`, `SetImageFloatingLayout`, and `RemoveImage`. Insert targets a paragraph
anchor and exact character boundary. Successful insert returns the new `ImageId`; list-returned
ids feed every later operation unchanged. Mutations are ordinary single undo steps. Valid no-ops
do not create history, and validation failures occur before snapshot creation.

## Writable subset and formats

Canonical embedded `w:drawing/wp:inline|wp:anchor/a:graphic/a:graphicData/pic:pic` pictures are
writable when the placement and surrounding inline boundary are safe. Legacy VML, external
linked images, multi-picture/non-canonical DrawingML, malformed structures, unsupported floating
geometry, and content-type/signature mismatches remain enumerable but read-only. Image mutations
under `TrackedChangeMode.RenderInline` return `tracked_operation_unsupported`; the supported
tracked-change vocabulary cannot faithfully represent a native picture edit.

DrawingML inside `mc:AlternateContent` is also read-only. Both its modern occurrence and any VML
fallback are enumerated, but changing only one compatibility branch would make consumers render
different images.

A picture whose `a:blip` carries an extension list is read-only for the same reason. An SVG
picture keeps its raster fallback in `a:blip/@r:embed` and the real art in an
`a:extLst/a:ext/asvg:svgBlip`; `Descendants(a:blip)` still counts one blip, so without this check
a replace would swap only the fallback — an SVG-aware renderer would keep displaying the OLD
image while the API reported success, and a remove could strip the fallback part and orphan the
SVG one. Writing both blips atomically is a feature, not a defect fix; until it exists the
occurrence reports `canMutate:false` with a reason naming the extension.

PNG, JPEG, GIF, BMP, and TIFF are insertable/replaceable. Input is capped at 64 MiB and rendered
width/height at 100,000 points. The parser validates format signatures and reads dimensions from
headers only; it does not fully decode pixels. Empty, truncated, malformed, unknown, or
content-type-mismatched input is rejected with a typed image error.

Existing WebP parts can be signature-inspected but are read-only. Open XML SDK 3.5.1 does not
expose a Word `ImagePartType` for WebP, so writable WebP would require inventing a package
capability the selected SDK surface does not provide. `GetImageCapabilities()` makes this
limitation explicit rather than inferring it from a failed insert.

## Units and floating layout

Rendered dimensions are points. When neither insert dimension is supplied, intrinsic pixels are
mapped at 96 DPI: one pixel is 0.75 point. Supplying one dimension with `PreserveAspect=true`
derives the other from intrinsic or current rendered aspect ratio. Both DrawingML extent copies
(`wp:extent` and `a:xfrm/a:ext`) are updated together.

`PreserveAspect` has exactly one meaning on `SetImageDimensions`: **scale the CURRENT rendered
box**. It never consults the media's intrinsic ratio. On `InsertImage` there is no current box
yet, so it scales the intrinsic size — the same rule applied to the only box that exists.

**`ReplaceImage` is deliberately dimension-preserving.** It rewrites `r:embed` and nothing else;
`wp:extent` and `a:xfrm/a:ext` keep the EMUs they had. The rendered box is a layout decision the
document author made, and silently resizing it on a byte swap would move body text. Replacing a
100×100 logo with a 4000×3000 photo therefore renders the photo squeezed into the old square
until the caller re-fits it. Re-fitting needs no new API: `ListImages()` re-reads the media on
every call, so immediately after the replace the occurrence already reports the NEW
`IntrinsicWidthPixels`/`IntrinsicHeightPixels`. Compute the box you want from those and write it
with `SetImageDimensions(id, width, height, preserveAspect: false)`, which takes an exact box.
Passing `preserveAspect: true` instead would scale the stale box and keep the old ratio.

Floating offsets, wrap distances, and relative positions are exact English Metric Units (EMUs),
not points. The writable layout subset supports `none` and `square` wrap, typed page/margin/
column/character horizontal references, page/margin/paragraph/line vertical references,
offset-or-alignment positioning, relative height, behind-document/lock/layout-in-cell/overlap
flags, and wrap side. Tight/through/top-and-bottom wrap, relative sizing, `simplePos`, duplicate or
mixed align/offset positions, malformed booleans/numerics, and unknown reference/alignment tokens
are reported with raw OOXML tokens and make the occurrence read-only. Position or wrap elements
with any unmodeled attributes or children are likewise preserved for inspection and rejected for
mutation rather than being silently replaced by the smaller modeled shape.

## Package topology, cleanup, and history

An image relationship belongs to the story part containing its markup. Inserts first reuse
identical content within that owner, then attach an identical package media part already used by
another story owner, and create a media part only when necessary. Equality includes both content
type and bytes. Drawing property ids are allocated document-wide, including headers and footers.

After image removal and generic destructive operations, owner-local image relationships are swept
only when the relationship id appears in **no attribute anywhere** in the owning part's XML.
Shared media remains until its final owning relationship is gone. Raw XML replacement performs
cleanup only after the replacement has validated successfully. `Save` repeats this safe sweep
over every story owner as the centralized package-normalization boundary, covering transforms
that do not use an image API; normalization does not create an undo entry.

The reference test is deliberately name-blind rather than a whitelist of `r:embed`/`r:link`/
`r:id`. Deletion is irreversible, and OOXML names image relationships through more attributes
than the DrawingML pair — VML and OLE spellings such as `o:relid` and `r:href` among them — so
enumerating the known ones destroys media referenced any other way. Matching on value alone is
safe because relationship ids are unique within a part: a non-reference attribute that happens
to hold the id keeps media alive, which is the recoverable direction. (`IM019` pins this
negative direction; `IM018` pins that a genuine orphan is still swept.)

**A render mutates the package.** `HtmlConversionOps.ConvertToHtml(session)` is implemented as
`session.Save(persistAnchorIds: true)`, so the sweep — and the rest of save-time normalization —
runs on a caller who only asked to look at the document. Nothing is lost today, because the sweep
now removes only provably unreferenced relationships, but the coupling is real: a session that
has been rendered is not byte-for-byte the session that was opened. Gating the sweep behind a
public setting was considered and deliberately deferred; it is a public-API decision, not a
defect fix.

Undo snapshots include image bytes/content types, exact media part URIs, every owner-local
embedded relationship id/target, and external `r:link` ids/targets. Restore rebuilds that layer at
the OPC package level and reopens the SDK graph, preserving topology across save/reopen,
undo/redo, shared owners, format replacement, and external links. Snapshot memory accounting
includes the captured media bytes.

## Transport surfaces

- .NET accepts `byte[]` and typed records directly.
- WASM/npm accepts `Uint8Array`; npm performs chunked base64 encoding at the JS/WASM boundary.
- Python accepts `bytes`; the stdio client encodes them as base64.
- JSON ops and MCP use an explicit `imageBase64` string. They do not fetch URLs or interpret file
  paths. MCP exposes the grouped `docxodus_images` tool.

MCP image mutators can also be used as `docxodus_mutations` steps; `capabilities` and `list` are
read-only and rejected there. Preview mode applies the same image operation and then restores its
snapshot, including the media-part and relationship layer.

The JSON shape is manually serialized/parsing-safe for trimming and uses snake-case enum tokens.
All clients expose the same versioned capabilities record and typed occurrence/layout models.
