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
zoom, and `PaginationOptions.scale`.

The browser producer inventories addressable source IDs before pagination moves content out of
staging. It separately inventories cited footnote definitions and the header/footer variants
selected by real pages. Publishing fails if any expected canonical source lacks a measurable page
fragment. Producers can mark deliberately non-rendered content with
`data-page-map-exclude="true"`; `hidden`, `aria-hidden="true"`, and inline hidden styles carry the
same meaning.

## Source identity and clones

`data-anchor` remains the bare Unid used by the editor, and an active bare value is unique across
the paginated DOM. If package stories reuse a bare Unid, only one legacy editor anchor remains
active; every source still retains its collision-safe canonical identity. Bare Unids are never
PageMap keys.

Every rendered addressable block instead carries `data-source-anchor-id="kind:scope:unid"`.
Paragraphs, headings, lists, tables, rows, cells, note definitions and their paragraphs, and
visible comment representations are covered. The converter rebuilds this identity map from its
final preprocessed `XElement` trees immediately before HTML transformation.

Pagination clones retain `data-source-anchor-id` and receive `data-page-number`,
`data-fragment-index`, and `data-page-fragment-id`. Continuations and repeated headers, footers,
and notes are presentation fragments: they do not duplicate the active `data-anchor`.

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
