# Native content controls

`DocxSession` treats Word structured-document tags (`w:sdt`) as first-class live
objects. `ListContentControls()` walks every requested story part and returns controls
in package story order and outer-before-inner document order. Every entry includes its
native `w:id`, type, placement, tag, alias, lock, placeholder state, binding facts,
owning part/scope, nesting parent/depth, current text/list values, and mutation status.

## Identity and projection

The public anchor is `sdt:{scope}:{unid}`. For a unique, valid signed 32-bit
`w:sdtPr/w:id`, `unid` is a deterministic hash of that native id; the scope identifies
the owning story. It therefore survives value edits and a normal clean save/reopen,
even though `PtOpenXml:Unid` bookkeeping is stripped. Missing, invalid, or duplicate
native ids remain enumerable under deterministic diagnostic anchors but are not
mutable. Repeating-item clones receive fresh native ids before their anchors are made
public.

`sdt` is an AnchorIndex kind in both the WML projector and the immutable IR emitter.
The IR captures projector-order anchor facts while its private package is open, so
index parity also holds with `RetainSources=false`. The wrapper remains transparent to
markdown, HTML, and `ListBlocks`; those surfaces continue to render/list its content.
`ListInlineSpans` additionally returns outer-to-inner `contentControlAnchorIds` for each
run.

## Mutations

The typed surface is deliberately operation-specific:

- `FillContentControlText` and `FillContentControlRichText`
- `SetContentControlChecked` and `SetContentControlDate`
- `SelectContentControlItem` for dropdowns and combo boxes
- `FillContentControlPicture`, using the native-image byte validation and relationship
  management shared with issue #453
- `AddRepeatingSectionItem` and `RemoveRepeatingSectionItem`

Fills preserve `w:sdt`, `w:sdtPr`, `w:sdtEndPr`, and metadata not owned by the
operation. Text fills retain representative run/paragraph properties and clear only
the showing-placeholder marker. Picture fills replace the image relationship without
rebuilding the wrapper. Repeating clones freshen every nested content-control id and
drawing `docPr` id, and reject clone-sensitive bookmark, comment, permission, custom
XML range, and note-reference markup. The final item cannot be removed.

Every successful operation is one undo/redo step. Whole-control fills are explicitly
rejected in `render_inline` tracked-change mode until issue #455 defines revision
semantics; surgical text/format operations inside a control remain available.

## Locks, bindings, and nesting

Content locks are effective through ancestors. A locked target or ancestor fails
without changing history. A whole-content replacement that would discard a nested
control is also refused; callers address the nested child directly.

Bindings fail closed by default. `bindingPolicy: "detach_target"` is the only opt-in:
it removes the selected control's own `w:dataBinding` before the mutation. It never
removes an ancestor binding and never edits or regenerates a Custom XML data part.
A target inside a bound ancestor is always refused.

## Transports

The shared JSON facade, WASM bridge, TypeScript package, Python host/client, and MCP
server expose the same typed operations. Options JSON is strict: the only fill option
is `bindingPolicy`, with `preserve` (default) or `detach_target`. The MCP grouped tool
is `docxodus_content_controls`; its mutating actions participate in
`docxodus_mutations` apply/preview rollback, while `list` is read-only and rejected as
a batch step. Picture bytes cross JSON transports only as base64.

Failures are structured `EditErrorCode` values, including not found, malformed,
unsupported family/placement, wrong type, locked, bound, invalid value, unsafe nested
fill, and repeating-section constraint errors. Refused operations do not consume undo
history.
