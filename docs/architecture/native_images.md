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
- `canMutate` plus `unsupportedReason`, so inspection never implies write support; and
- `operations`, the per-occurrence matrix (issue #762): one `{operation, canMutate, reason}` row
  for each of `replace`, `embed_linked`, `set_dimensions`, `set_metadata`, `set_floating_layout`
  and `remove`, answered for the session's current tracked-change mode.

`canMutate` summarizes the standard set — replace, set_dimensions, set_metadata, remove, and
set_floating_layout for a floating picture — and `unsupportedReason` names the first blocker.
The matrix is the finer truth: a linked picture reports `canMutate:false` because replace is
refused, while its `set_metadata`, `set_dimensions`, `set_floating_layout`, `remove` and
`embed_linked` rows are writable. Discovery and execution share one gate, so a row that says
`canMutate:true` is a call that will run, and a refused row carries the reason the call would
return.

One non-canonical drawing can contain several blips. Each is returned separately with `:subN`
on the common drawing id. These rows are read-only. A drawing with no identifiable blip is also
reported as unsupported instead of disappearing — that guarantee is scoped to drawings the
projection can address, i.e. those inside a `w:p` that resolves to an anchor. A `w:drawing`
whose nearest paragraph has no anchor is skipped entirely and appears in no listing. IDs are
structural occurrence IDs, not media part IDs: two occurrences may deliberately share one
media part.

The mutation surface is `InsertImage`, `ReplaceImage`, `EmbedLinkedImage`, `SetImageDimensions`,
`SetImageMetadata`, `SetImageFloatingLayout`, and `RemoveImage`. Insert targets a paragraph
anchor and exact character boundary. Successful insert returns the new `ImageId`; list-returned
ids feed every later operation unchanged (except a tracked edit, which returns the id of the
inserted copy — see below). Mutations are ordinary single undo steps. Valid no-ops do not create
history, and validation failures occur before snapshot creation.

## Coverage matrix (issue #762)

`GetImageCapabilities().Markups` publishes the same matrix per markup family that each listed
occurrence answers for itself. Every family shares the writers — one relationship re-pointer, one
extent writer, one metadata writer, one removal — applied to the occurrence's outer element
(`w:drawing`, `w:pict`, or the whole `mc:AlternateContent`), so adding coverage did not add a
second code path per operation.

| Markup | Writable | Refused, and why |
|---|---|---|
| `embedded_picture` — canonical `w:drawing/wp:inline\|wp:anchor/…/pic:pic` | replace, set_dimensions, set_metadata, set_floating_layout (floating only), remove | `embed_linked`: the picture is already embedded |
| `linked_picture` — `a:blip/@r:link` (with or without a stale `r:embed`) | embed_linked, set_dimensions, set_metadata, set_floating_layout, remove | `replace`: external media is never converted in place. `EmbedLinkedImage(id, bytes)` is the explicit conversion — the caller fetches the bytes (the session never does), they become an owner-local media part, `r:link` is dropped, and the external relationship is swept with its last reference. |
| `extended_picture` — an `a:blip` extension carrying its own relationship (SVG `asvg:svgBlip`, an artistic-effect `a14:imgLayer`) | set_dimensions, set_metadata, set_floating_layout, remove | `replace`/`embed_linked`: they would swap only the raster fallback, so an SVG-aware renderer would keep showing the old art. Removal orphans both payloads, and the name-blind sweep drops both parts. |
| `legacy_vml` — `w:pict/v:shape/v:imagedata` with an `r:id` | replace, set_dimensions (the shape's CSS `style` width/height, rewritten in place), set_metadata (`v:shape/@alt`, `v:imagedata/@o:title`), remove | `set_floating_layout`: VML positioning is a different model and is not written. A shape without a parsable width/height style refuses set_dimensions. |
| `alternate_content` — a run holding one `mc:Choice` with exactly this drawing and at most one `mc:Fallback` holding one VML picture | replace (when every branch references one embedded relationship — one re-point changes them all), set_dimensions and set_metadata (written to every branch), remove (the whole `mc:AlternateContent`) | `set_floating_layout`: the branches carry independent geometry. The fallback occurrence is inspection-only and says to operate on the modern one. |
| `multi_picture` — several blips in one drawing (`:subN` ids), groups, no identifiable blip, missing `wp:docPr`/`wp:extent`/`a:xfrm` | nothing | structural reason |

Overlays refine a row for one occurrence: an inline picture refuses set_floating_layout; a
floating picture whose layout holds unmodeled tokens refuses only set_floating_layout (its
`FloatingLayout` keeps the raw tokens for inspection; replace, sizing, metadata and removal still
work because they never rewrite the anchor); a picture inside a tracked deletion or move, a
simple-field result, or a paragraph with a complex field refuses everything; and under
`render_inline` a picture inside another author's insertion refuses everything, because editing
it would rewrite their revision. Hyperlinks, smart tags and content-control wrappers hold runs
and revision envelopes alike, so a picture inside them is addressable.

Formats: PNG, JPEG, GIF, BMP, TIFF and WebP are insertable and replaceable. WebP lands as an
`image/webp` media part with a `.webp` target through the SDK's `PartTypeInfo` — the same
(content type, extension) pair the built-in `ImagePartType` table entries are — rather than an
invented package capability; the capability row notes that consumers without WebP support show a
placeholder. Input is capped at 64 MiB and rendered width/height at 100,000 points. The parser
validates format signatures and reads dimensions from headers only; it does not fully decode
pixels. Empty, truncated, malformed, unknown, or content-type-mismatched input is rejected with a
typed image error.

## Tracked mutations (issue #762)

Under `TrackedChangeMode.RenderInline` every image operation is recorded natively. A new picture
is inserted inside `w:ins`. A change to an existing picture is recorded the one way
WordprocessingML can express it: the picture's run (isolated first, so text sharing the run moves
to sibling runs with the same properties) becomes a `w:del`, and a re-identified copy carrying the
change — fresh Unids, fresh `wp:docPr` ids — becomes a `w:ins`; remove records the deletion
alone. The operation returns the inserted copy's id. Listing then shows both: the deleted
occurrence is inspectable but refuses everything (`image is inside a tracked deletion`), the
inserted one is live.

This holds for sizing, metadata and layout too, not only bytes: Word itself does not mark a
picture resize as a revision, but recording it as a swap is what makes the acceptance criteria
true — accepting yields exactly the intended picture, and rejecting restores the original bytes,
relationship, metadata and geometry because the original run was never edited. The media part
the deleted run references stays in the package until the deletion is accepted; the orphan
sweep that revision resolution already runs then drops whichever part lost its last reference.

A picture that already sits inside the session author's own `w:ins` is edited in place, so the
replace-then-refit recipe below stays one revision pair instead of nesting a second. Tracked
inserts into a paragraph that already carries revisions land normally; only a boundary strictly
inside another inline container is refused, as for text.

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
not points. The writable layout subset supports every wrap form — `none`, `square`, `tight`,
`through` and `top_and_bottom` — typed page/margin/column/character horizontal references,
page/margin/paragraph/line vertical references, offset-or-alignment positioning, relative height,
behind-document/lock/layout-in-cell/overlap flags, and wrap side. Tight and through wrap carry a
`WrapPolygon`: the outline text follows, as a start vertex plus at least two line segments in
DrawingML's 21600-unit picture space, with `Edited` mirroring `wp:wrapPolygon/@edited`. Word
derives that outline from picture transparency; the session has no decoder, so a layout written
without vertices gets the picture rectangle (`ImageWrapPolygon.Rectangle`, which is also what
Word writes for an opaque picture) and a layout read from the document always carries the
outline it holds. Relative sizing, `simplePos`, duplicate or mixed align/offset positions,
malformed booleans/numerics, and unknown reference/alignment tokens are reported with raw OOXML
tokens and make `set_floating_layout` read-only. Position or wrap elements with any unmodeled
attributes or children — distances or effect extents on the wrap element itself, a polygon that
is not a plain start/lineTo list — are likewise preserved for inspection and rejected for
mutation rather than being silently replaced by the smaller modeled shape.

## Package topology, cleanup, and history

An image relationship belongs to the story part containing its markup. Inserts first reuse
identical content within that owner, then attach an identical package media part already used by
another story owner, and create a media part only when necessary. Equality includes both content
type and bytes. Drawing property ids are allocated document-wide, including headers and footers.

After image removal and generic destructive operations, owner-local image relationships are swept
only when the relationship id appears in **no attribute anywhere** in the owning part's XML.
Shared media remains until its final owning relationship is gone. Raw XML replacement performs
cleanup only after the replacement has validated successfully.

**The sweep's boundary is mutation, not serialization.** Orphaning is something a mutation does;
`DocxSession.InvalidateProjectionCache` — the single point every op reaches once its edit has
landed in the live XML — runs the package-wide sweep over every story owner. That covers the
transforms which drop a `w:drawing` without any image API involved (`DeleteBlock`, `DeleteRange`,
table row/column deletes, `ReplaceText`, the raw XML ops). Most of those additionally sweep their
own resolved owner; the package-wide pass is what makes the invariant structural rather than a
per-op checklist, and it is the only thing that covers an op whose edit lands in a part other
than the one it resolved. Normalization does not create an undo entry. The undo/redo restore
paths deliberately skip the sweep: a snapshot is authoritative over relationship topology.

The cost is bounded below what every op already pays. `SweepOrphanedImages` returns before
reading any XML when the owner holds no image relationship, so an image-free story is free and
the sweep is never what materializes a part's `XDocument`; when there are relationships it
resolves the whole candidate set in **one** attribute walk rather than one walk per relationship,
so the per-mutation cost does not grow with the image count. That walk is strictly cheaper than
the `TakeSnapshot` each mutating op already runs, which *serializes* the very trees the sweep
merely reads.

The reference test is deliberately name-blind rather than a whitelist of `r:embed`/`r:link`/
`r:id`. Deletion is irreversible, and OOXML names image relationships through more attributes
than the DrawingML pair — VML and OLE spellings such as `o:relid` and `r:href` among them — so
enumerating the known ones destroys media referenced any other way. Matching on value alone is
safe because relationship ids are unique within a part: a non-reference attribute that happens
to hold the id keeps media alive, which is the recoverable direction. (`IM019` pins this
negative direction, across a save, a render, and a mutation; `IM018` pins that a genuine orphan
is still swept — by a mutation whose own owner is a different story part.)

**A render does not mutate the package.** `HtmlConversionOps.ConvertToHtml(session)` is
implemented as `session.Save(persistAnchorIds: true)`, so anything `Save` normalized would run on
a caller who only asked to look at the document. `Save` — and therefore every render — is now
read-only with respect to relationships and media: an orphan present in the opened bytes is still
there after any number of renders and saves, and disappears only when the session is next
mutated. `IM027` pins that topology and the media payloads are unchanged across repeated
`ConvertToHtml` and both save flavours.

One consequence worth stating: a pre-existing orphan in an input document is no longer cleaned up
by saving it back unchanged. That is deliberate — open/save is lossless, and the session does not
silently delete media it never touched.

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

MCP image mutators — `embed_linked` included — can also be used as `docxodus_mutations` steps;
`capabilities` and `list` are read-only and rejected there. Preview mode applies the same image
operation and then restores its snapshot, including the media-part and relationship layer and,
under `render_inline`, the revision markup the operation wrote.

The JSON shape is manually serialized/parsing-safe for trimming and uses snake-case enum tokens.
All clients expose the same versioned capabilities record and typed occurrence/layout models.
