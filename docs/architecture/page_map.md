# Portable PageMap and page citations

`PageMap` is the versioned handoff between a renderer that has performed physical pagination and
the stateful `DocxSession` APIs that answer search and citation requests. It deliberately does not
put pagination into core Docxodus: the browser `PaginationEngine` is currently the page authority,
and issue #434 remains the dependency for a server-side paginated HTML substrate.

## Authority and availability

A renderer may publish an available map only after it has created fixed page boxes and measured
every rendered, addressable source node. Continuous HTML has no physical page substrate and must
use an unavailable map:

```json
{
  "schemaVersion": 1,
  "mode": "continuous",
  "availability": "unavailable",
  "documentVersion": 12,
  "rendererFingerprint": "continuous-preview-v1",
  "pages": [],
  "fragments": []
}
```

No surface derives page numbers from document length, section metadata, or
`EstimatedPageCount`. That metadata field is explicitly labeled `heuristic`; it is not a citation
source.

## Version 1 contract

An available paginated map contains:

- the exact `documentVersion` from the session that produced the rendered HTML;
- a caller-defined `rendererFingerprint` identifying every layout-affecting renderer input;
- document-ordered physical pages, including global `pageNumber`, `pageInSection`, dimensions in
  points, optional non-negative `sectionIndex`, and stable `pageName`;
- one fragment for every visible piece of every rendered source anchor, with a page-qualified
  `fragmentId`, full canonical `kind:scope:unid` `anchorId`, contiguous `fragmentIndex`, page-relative
  point geometry, owning story (`body`, `header`, `footer`, `footnote`, `endnote`, or `comment`), and
  table-cell ownership.

Geometry is normalized against the known page box, so it is independent of CSS pixels, browser
zoom, and `PaginationOptions.scale`. With `commentColumnPlacement: "balloons"` (Word's All Markup
layout, `all_markup_layout.md`) the page box stays paper-sized while its content is shrunk beside a
markup area, so fragments record where the shrunken content prints on the paper.

The browser producer inventories addressable source IDs before pagination moves content out of
staging. It separately inventories cited footnote definitions and the header/footer variants
selected by real pages. Publishing fails if any expected canonical source lacks a measurable page
fragment. Producers can mark deliberately non-rendered content with
`data-page-map-exclude="true"`; `hidden`, `aria-hidden="true"`, and inline hidden styles carry the
same meaning. Content an exact-height table row (`w:hRule="exact"`) hides whole is not missing
either: Word hides what such a row cannot hold, so a source whose layout box lies wholly outside its
row cell's clip box (`data-docx-exact-row-cell`) has no fragment and does not fail publishing
(issue #1168), and a citation of it reads `anchor_not_mapped`. Only that clip counts; content the
body band or a pagination clone clips away whole still fails, because there it was lost.

A source's fragment is where it renders, measured by `measureRenderedSource` in
`npm/src/pagination.ts` — the one function fragment-identity normalization, PageMap materialization
and the standalone export's offline re-verification all share. Normally that is the element's own
box, clipped by the page and by every overflow-clipping ancestor. A paragraph whose only content is
a floating text box or shape is different: the paginator promotes the drawing out of the text column
into the page box (recording its host in `data-docx-anchor-host`), retaining the paragraph-mark line
independently of the drawing's extent. Such a paragraph measures as the union of what it renders — its own descendants
and the drawings promoted out of it, each clipped along its own ancestor chain — so it is not
reported as missing (issue #849).

Floating shape clearance is separate from the promoted ink: top/bottom exclusions stay in the text
flow while the SVG moves to its page position. The page uses `isolation: isolate` so background
drawings remain above its paper when standalone export removes `contain` and `will-change`.
PageMap geometry alone cannot prove that a drawing is painted above an opaque background; the
floating-shape export regressions also inspect PDF paint order (#1346).
When a paragraph splits beside a float, its candidate head is measured with the preceding page
content. Measuring it alone would admit more text than its narrowed first lines can hold. The tail
is measured without the previous page's float, so every glyph remains within its own body band.

Read-only `paginateHtml()` defaults to a scrollable comment panel outside each page. The source
comment subtree and its outside presentation, like the `"balloons"` placement's markup area, balloons
and leaders, carry `data-page-map-exclude="true"`: they remain
readable UI, but have no physical page substrate. Body and running-story coverage is still required.
The lower-level `PaginationEngine` and standalone export retain physical-margin placement; callers
can also select it explicitly with `commentColumnPlacement: "margin"` when comment page citations
are required. Width and gap options affect the outside panel, and physical page dimensions remain
unchanged.

Standalone physical-margin export appends columns that cannot print in full (#1490). These
comments remain ordinary PageMap sources: each projected paragraph retains its `p:cmt` identity,
and inner spans carry its comment and ancestor-thread identities through every Range split.
Hiding a metadata header therefore cannot hide the comment's owner from the map. The added
presentation section has the last paper geometry and no running stories; it does not change the
DOCX section model. Interactive scrolling remains the default `commentOverflow: "scroll"` policy;
the exporter selects `"appendix"`. Appendix paragraphs always permit fragmentation, independently
of the body paragraph policy. Projected reply headers retain ancestor owners around their flex
layout, including when a header occupies a page by itself.

Endnote-style comments use the same paragraph-group splitter as bordered paragraphs (#1505).
The converter's section, list, item and body wrappers remain around each measured candidate,
so reply indentation, metadata, links and ancestor comment identities survive continuation.
The title and comment headers keep with the following content. Only the known converter shape
in ordinary block layout is admitted; custom containers retain the whole-block fallback.
Splits use the current section's actual body band, including continuous sections and tall headers,
and obey the normal paragraph-fragmentation option and export resource limits.

## Source identity and clones

`data-anchor` remains the bare Unid used by the editor, and an active bare value is unique across
the paginated DOM. If package stories reuse a bare Unid, only one legacy editor anchor remains
active; every source still retains its collision-safe canonical identity. Bare Unids are never
PageMap keys.

Every rendered addressable block instead carries `data-source-anchor-id="kind:scope:unid"`.
Paragraphs, headings, lists, tables, rows, cells, note definitions and their paragraphs, and
visible comment representations are covered. The converter indexes canonical identities before
preprocessing and resolves rewritten elements by their retained part URI and Unid.

Textbox hosts also carry `data-docx-required-source-anchors`, independently derived from the
selected OOXML textbox paragraphs. Pagination includes those IDs even when HTML conversion
omits a drawing member (#1482). Only the selected compatibility branch contributes. An image
or chart alternative suppresses its own text story only when that representation actually
renders; callback or missing-image fallbacks retain the text expectations. The source inventory
is collected after projection selects representations, without deriving it from projected HTML.
As with ordinary anchors, unselected running-story variants do not become required.

Pagination clones retain `data-source-anchor-id` and receive `data-page-number`,
`data-fragment-index`, and `data-page-fragment-id`. Continuations and repeated headers, footers,
and notes are presentation fragments: they do not duplicate the active `data-anchor`. So are the
copies of a table's header rows (`w:tblHeader`) that open each page a split table continues on
(`data-docx-repeated-header-row`, issue #1171): the row has one fragment per page it opens.

Bordered footnote paragraphs use the body's shared border-group fragmenter (#1491), with each
candidate measured in the actual note payload. Both the citation page and later continuation
pages and nested table cells use that path. Range fragments retain the paragraph's source identity; only the group's
top/bottom border at an internal page break is omitted. The footnote path remains independent
of the body paragraph-fragmentation option. Explicit line keeps retain the existing fresh-band
retry policy; a kept heading stays with a nonempty prefix of the following paragraph. Indivisible
overflow still fails export.

Margin-note copies strip duplicate element IDs. Their generated return links are resolved only
after all pages exist, against one actual rendered reading-content reference. Repeated notes can
return to the same reference without adding another bookmark target or editor anchor. Link
eligibility uses painted geometry; a visible citation can still be excluded from the authored
source inventory.

## Registration and invalidation

Register a map through `DocxSession.RegisterPageMap` (`registerPageMap` in TypeScript,
`register_page_map` in Python, or `docxodus_pagination` over MCP). Registration validates the schema,
enum discriminators, exact document version and optional expected fingerprint, page/section order,
canonical live anchors, story/scope and table ancestry, page bounds, and fragment chronology.

A successful mutation, undo, or redo advances the session version and immediately makes the map
stale. Citation reads require an exact `{documentVersion, rendererFingerprint}` token. Their
unavailable result is typed as one of `no_page_map`, `continuous_mode`,
`stale_document_version`, `renderer_fingerprint_mismatch`, or `anchor_not_mapped`.

## Browser workflow

```ts
const version = session.getVersion();
const rendererFingerprint = 'chromium-140|fonts-v3|docxodus-pagination-v1';

const result = paginateHtml(html, viewer, {
  layoutToken: { documentVersion: version, rendererFingerprint },
});
const registration = session.registerPageMap(result.pageMap!, rendererFingerprint);
if (!registration.success) throw new Error(registration.message);

const citation = session.getPageCitation(anchorId, {
  documentVersion: version,
  rendererFingerprint,
});
navigateToPageCitation(viewer, citation);
```

`PaginatedDocument` accepts the same `layoutToken` and an optional `citation`; it navigates and
highlights after pagination. Search and scoped projection APIs accept the exact citation token and
attach citation envelopes without changing their results when citations are not requested.

The current MCP inline widget renders continuous HTML, so it can display the cited page label and
highlight an anchor but reports `pageNavigation: "unavailable_continuous_preview"`. Once #434
provides a physical server-rendered substrate, the same PageMap and navigation contract can consume
it without inventing a second page model.
