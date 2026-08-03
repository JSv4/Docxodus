# Round-three MCP comparison results

Verified 2026-08-02 on branch `smoke-test-mcp-round-three` using the requested October 2025 DOCX, the replacement server built for .NET 10, and LibreOffice 25.8.7.3 for independent PDF rendering.

## Outcome

The mapped workflows are semantically equivalent for the exercised operations. The replacement run completed 56/56 calls and eight embedded assertions with zero failures; the reference run completed 64/64 calls with zero final workflow failures. A separate close/reopen pass completed 9/9 calls and nine assertions, confirming all expected markers, one created table, two resolved threaded comments, and zero remaining revisions.

Both outputs render as 49 US-letter pages. Their extracted text is identical apart from whitespace placement in the final table region. The final page has the same visible text, inline formatting, Roman-list hierarchy, header shading, borders, row-height behavior, and resolved-review content.

| Observable | Reference output | Replacement output |
| --- | ---: | ---: |
| Expected marker occurrences | 9 | 9 |
| Preview/rejected marker occurrences | 0 | 0 |
| Comments / replies / resolved thread entries | 2 / 1 / 2 | 2 / 1 / 2 |
| Remaining insertions / deletions | 0 / 0 | 0 / 0 |
| Added paragraphs / numbering properties | 18 / 4 | 18 / 4 |
| Table rows / cells | 3 / 9 | 3 / 9 |
| Table widths, twips | 3000 / 3795 / 2400 | 3000 / 3800 / 2400 |
| Repeat header / no split / minimum height | yes / yes / 480 | yes / yes / 480 |
| PDF pages | 49 | 49 |
| DOCX file parts | 46 | 47 |
| DOCX compressed bytes | 146,686 | 218,417 |
| DOCX uncompressed bytes | 1,113,421 | 1,101,694 |

The 5-twip middle-column variance is 0.25 point and is below visible significance. The replacement writes the requested 190-point width exactly; the reference output rounds it down.

## Material differences

1. The servers are semantic substitutes, not wire-protocol substitutes. Tool names, grouping, argument shapes, result shapes, and call counts differ. An existing client needs an adapter; simply changing the server command is not sufficient.

2. Surgical range replacement under the replacement server's tracked-change mode currently writes direct text instead of `w:del`/`w:ins`. The workflow uses full-block tracked replacement for accept/reject coverage. This is material for clients that require tracked edits within only part of a paragraph.

3. The reference surface advertised tracking on some direct replacement, structural-create, and comment paths that did not behave consistently at runtime: direct calls rejected the tracking argument in some cases, while tracked structural creation did not produce revisions. The completed reference workflow therefore used its mutation-batch path for tracked rewrites.

4. Non-body story addressing differs. Reference story locators required opaque story identifiers not supplied by ordinary discovery, and an incompletely identified footnote request fell back to the body. The replacement exposes named header/footer/footnote scopes directly. The smoke workflow inventories all stories but limits mutations to the body to keep the cross-server workload deterministic.

5. Package rewrite strategies differ materially. Both outputs changed 24 of 44 original parts by byte hash, but the affected parts are different. The reference output rewrote custom XML, core/custom properties, styles, settings, theme, font table, web settings, and expanded `footnotes.xml` by 8,141 bytes. The replacement preserved those parts byte-for-byte while its projected header/footer/endnote/footnote XML parts each changed by only three serialization bytes.

6. The reference output coalesced field-instruction nodes from 201 to 135, while the replacement retained the source's 201 nodes. Both retain 405 field characters and the canonical field-instruction token stream has the same SHA-256 across source and both outputs, so this is normalization rather than a field-semantic difference.

7. The replacement package is 48.9% larger on disk despite being smaller uncompressed. This is a ZIP compression-efficiency difference, not added document content. It is material for storage/bandwidth-sensitive use.

8. Thread metadata differs by one part: the replacement writes `commentsIds.xml` in addition to `comments.xml` and `commentsExtended.xml`; the reference output writes the latter two. Both persist the same root/reply relationship and resolved state.

## Conformance fixes produced by the smoke test

The comparison exposed and fixed four replacement-side gaps:

- inserted headings now write an explicit no-numbering override so legally numbered source styles cannot prefix new headings;
- list removal now works for heading/list-item anchors and style-inherited numbering;
- resolving or reopening a root comment propagates through its reply subtree;
- table row options now support repeat-header state, page-split prevention, and auto/at-least/exact row heights.

Focused regressions cover each behavior. The final selections passed 5/5 list/heading tests and 98/98 MCP-dispatcher, table-edit, and comment-authoring tests; the complete replacement workflow also passes after rebuilding and reopening the saved output.

## Evidence snapshot

- Source SHA-256: `d75600769c12724990de48149d7a2bb161f3522daa54b1783672f93697d87d29`
- Reference output SHA-256: `3b9fddcdca57656597a41c2590c1b73b5f4cd2bd3f619084d4c8c66440d375e4`
- Replacement output SHA-256: `6202717691c9837290a51a51aa9c9446e1fd6a18c495b045ab67e47af2d49d64`
- Replacement trace: `/tmp/docxodus-mcp-round-three/local-trace.json`
- Reopen-validation trace: `/tmp/docxodus-mcp-round-three/local-validation-trace.json`
- Package comparison: `/tmp/docxodus-mcp-round-three/comparison.json`
- Rendered outputs: `/tmp/docxodus-mcp-round-three/render/reference/reference.pdf` and `/tmp/docxodus-mcp-round-three/render/local-final4/local.pdf`

The replacement hash changes across reruns because newly authored comment/thread identifiers are generated per session; the semantic and package-shape assertions remain stable.
