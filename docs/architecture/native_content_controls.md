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
even though `PtOpenXml:Unid` bookkeeping is stripped. Missing, invalid, or
package-wide duplicate native ids remain enumerable under deterministic diagnostic
anchors but are not mutable. Repeating-item clones receive fresh native ids before
their anchors are made public.

Mutation also requires an exact SDT envelope: one `w:sdtPr`, one `w:sdtContent`,
one `w:id`, no more than one mutually exclusive family marker, and at most one
`w:lock` carrying a native lock value. The same malformed-envelope gate applies through
ancestors, so a valid child cannot bypass a malformed outer lock. Malformed controls stay
enumerable under diagnostic anchors and fail before an undo snapshot. A repeating template
is cloneable only when every nested SDT satisfies the same invariant.

`sdt` is an AnchorIndex kind in both the WML projector and the immutable IR emitter.
The IR captures projector-order anchor facts while its private package is open, so
index parity also holds with `RetainSources=false`. The current Markdown oracle indexes
the outer SDT but omits content delivered through an inline or block SDT, so adding the
anchor does not change historical Markdown bytes. HTML and `ListBlocks` remain wrapper-
transparent and flatten/render the contained blocks.
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
the showing-placeholder marker. Empty rich-text input normalizes to one schema-safe
empty paragraph or run payload, preserving the wrapper and placeholder definition while
clearing its showing-state marker. Picture fills replace the image relationship without
rebuilding the wrapper. Discovery and fill share the same picture topology gate: exactly
one canonical, embedded, mutable image must belong to the control, so `CanMutate` cannot
advertise zero-image, multi-image, linked, or unsupported picture controls as writable.
Repeating clones freshen every nested content-control id and
drawing `docPr` id, and reject clone-sensitive bookmark, comment, permission, custom
XML container/range, move, note-reference, and `w14:paraId`/`w14:textId` markup. The
clone gate also rejects every live tracked-revision carrier recognized by the revision
registry; duplicating such markup would duplicate its native revision ids and make later
resolution ambiguous. The final item cannot be removed.

Dropdown selection writes the selected item's native `w:lastValue` as well as its
displayed text. Combo boxes do the same for a listed item and also accept custom text;
custom text becomes both the displayed payload and `w:lastValue`. Checkbox fills honor
the selected `w14:checkedState`/`w14:uncheckedState` font on the produced glyph run.

Whole-content replacement and repeating-item removal use the same bookmark-removal
gate as other structural edits: crossing or externally referenced ranges fail before
history changes. Rich-text links are validated against the document that will remain
after replacement, so a payload cannot target a bookmark it simultaneously removes.
After a successful replacement/removal, owner-local hyperlink relationships are
promoted or reference-counted away and unreferenced image parts are swept. Undo/redo
restores that XML and package relationship topology together.

Every successful operation is one undo/redo step. Text, rich-text, checkbox, date,
dropdown, and combo-box fills support inline and block controls only. Row/cell controls
remain enumerable; picture and repeating-section operations use their own structural
shape checks instead of the text-placement rule. Whole-control fills are rejected in
`render_inline` tracked-change mode because they do not yet have a faithful replacement
revision encoding; surgical text/format operations inside a control remain available.
Empty and nested-SDT-only payloads derive row/cell/block/inline placement from their nearest
owning content-model boundary rather than defaulting to block placement.

## Anchor and receipt lifecycle

Typed fills preserve the selected wrapper identity and return that `sdt` anchor in
`Modified`. `AddRepeatingSectionItem` returns the fresh item anchor in `Created` and the
section anchor in `Modified`. `RemoveRepeatingSectionItem` returns the item anchor in
`Removed` and the section anchor in `Modified`. These identities remain usable through
undo/redo according to whether the corresponding wrapper is live.

Generic tracked `DeleteRange`/`DeleteSection` keeps a selected SDT wrapper live until
revision resolution. Its receipt therefore reports the wrapper `sdt` anchor and every
retained descendant anchor in `Modified`; a structural fall-through that is actually
removed appears in `Removed`.

## Locks, bindings, and nesting

Content locks are effective through ancestors. A locked target or ancestor fails
without changing history. A whole-content replacement that would discard a nested
control is also refused; callers address the nested child directly.

For repeating sections, `CanMutate` describes the default operation honestly: a
section must have a safe final clone template, and an item is removable only when it
is a direct child, at least one sibling item will remain, and its wrapper is not
locked. Orphaned/non-direct items and clone-sensitive templates remain enumerable with
the corresponding diagnostic.

Bindings fail closed by default. Both `w:dataBinding` and the Office 2013
`w15:dataBinding` form are recognized. `bindingPolicy: "detach_target"` is the only
opt-in: it removes the selected control's own native binding element before the
mutation. It never removes an ancestor binding and never edits or regenerates a
Custom XML data part. A target inside a bound ancestor is always refused.

## Transports

The shared JSON facade, WASM bridge, TypeScript package, Python host/client, and MCP
server expose the same typed operations. Options JSON is strict: the only fill option
is `bindingPolicy`, with `preserve` (default) or `detach_target`. The MCP grouped tool
is `docxodus_content_controls`; it advertises the same optional optimistic
`preconditions` guard as the other mutating tools, and its mutating actions participate in
`docxodus_mutations` apply/preview rollback, while `list` is read-only and rejected as
a batch step. Picture bytes cross JSON transports only as base64.
An omitted date `displayText` selects the invariant default; an explicitly empty
string remains empty through the JSON, WASM, and TypeScript layers.

Failures are structured `EditErrorCode` values, including not found, malformed,
unsupported family/placement, wrong type, locked, bound, invalid value, unsafe nested
fill, and repeating-section constraint errors. Refused operations do not consume undo
history.
