#nullable enable

using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.RegularExpressions;
using Docxodus;
using Docxodus.Internal;

namespace DocxodusWasm;

/// <summary>
/// JSExport bridge for <see cref="DocxSession"/>. Sessions live on the .NET heap
/// and persist across JSExport calls — keyed by an integer handle returned from
/// <see cref="OpenSession"/>. JS-side code must call <see cref="CloseSession"/>
/// when done; sessions are not eligible for GC otherwise.
///
/// All wire-format work — JSON serialization, settings/format-op parsing, the
/// handle pool — lives in <see cref="DocxSessionOps"/> / <see cref="DocxSessionJson"/>
/// / <see cref="SessionRegistry"/>. This file is a thin JSExport-attributed
/// shell so the WASM and stdio NDJSON transports stay byte-for-byte identical.
/// </summary>
[SupportedOSPlatform("browser")]
public static partial class DocxSessionBridge
{
    [JSExport]
    public static int OpenSession(byte[] bytes, string settingsJson) =>
        DocxSessionOps.OpenSession(bytes, DocxSessionJson.ParseSettings(settingsJson));

    [JSExport]
    public static void CloseSession(int handle) => DocxSessionOps.CloseSession(handle);

    /// <summary>Mint a complete blank DOCX (a "New document" seed) as bytes.</summary>
    [JSExport]
    public static byte[] CreateBlankDocx() => DocxSessionOps.CreateBlankDocx();

    [JSExport]
    public static string Project(int handle) => DocxSessionOps.Project(handle);

    /// <summary>
    /// Ordered top-level render units per scope container (body / footnotes /
    /// endnotes), as JSON: <c>{"body":[{"id","kind"},…],"footnotes":[…],"endnotes":[…]}</c>.
    /// A table is ONE body unit; a note definition is one notes unit. The editor's
    /// incremental reconciler diffs its DOM against this after a structural op.
    /// </summary>
    [JSExport]
    public static string ListBlocks(int handle) => DocxSessionOps.ListBlocks(handle);

    /// <summary>
    /// Citation-ordered footnotes (or endnotes) as JSON
    /// <c>[{"id","defAnchorId","ordinal"},…]</c> — the id↔ordinal authority the
    /// editor's reconciler walks when it renumbers rendered note chrome (marker
    /// sup text, hrefs, list values) after a note insert/delete.
    /// </summary>
    [JSExport]
    public static string ListNotes(int handle, bool endnotes) => DocxSessionOps.ListNotes(handle, endnotes);

    /// <summary>
    /// The anchor index alone as <c>{"anchorIndex":{…}}</c> — the editor's per-op
    /// anchor-map refresh, without marshaling the whole markdown projection.
    /// </summary>
    [JSExport]
    public static string ListAnchors(int handle) => DocxSessionOps.ListAnchors(handle);

    /// <summary>
    /// Batch block render: <paramref name="anchorIdsJson"/> is a JSON string array of
    /// anchor ids; returns a JSON object mapping each id to its HTML element (null for
    /// an id that failed to resolve). One throwaway document and one converter run for
    /// the whole batch, with real sibling context and true list-marker numbers.
    /// </summary>
    [JSExport]
    public static string RenderBlocksHtml(int handle, string anchorIdsJson, string cssPrefix, bool fabricateClasses)
    {
        // NOTE: success output is itself a JSON object, so the "HTML never starts with
        // '{'" error convention of RenderBlockHtml does not apply here — callers parse
        // and check for an "error" property (anchor-id keys always contain a colon, so
        // they can never collide with it).
        try { return DocxSessionOps.RenderBlocksHtml(handle, anchorIdsJson, cssPrefix, fabricateClasses); }
        catch (System.Exception ex) { return $"{{\"error\":\"{JsonEncodedText.Encode(ex.Message ?? string.Empty)}\"}}"; }
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.ProjectAnchor"/>. <paramref name="depth"/>
    /// uses the numeric layout of <see cref="ProjectionDepth"/> (SelfOnly=0,
    /// Subtree=1, SubtreeAndFollowingSiblings=2). Returns a JSON object with
    /// the standard MarkdownProjection shape (markdown + anchorIndex).
    /// </summary>
    [JSExport]
    public static string ProjectAnchor(int h, string anchorId, int depth) =>
        DocxSessionOps.ProjectAnchor(h, anchorId, (ProjectionDepth)depth);

    /// <summary>
    /// Render a single block from the live session to faithful HTML — the editor's
    /// incremental per-block re-render. Returns the block's HTML element, or a JSON
    /// <c>{"error": "..."}</c> object on failure (rendered HTML always begins with
    /// '&lt;', so the leading character disambiguates success from error).
    /// </summary>
    [JSExport]
    public static string RenderBlockHtml(int h, string anchorId, string cssPrefix, bool fabricateClasses)
    {
        try { return DocxSessionOps.RenderBlockHtml(h, anchorId, cssPrefix, fabricateClasses); }
        // Reflection-free error JSON: the trimmed WASM build disables reflection-based
        // JsonSerializer, so serializing an anonymous type here would itself throw
        // (JsonSerializerIsReflectionDisabled) and mask the real failure as an uncaught crash.
        catch (System.Exception ex) { return $"{{\"error\":\"{JsonEncodedText.Encode(ex.Message ?? string.Empty)}\"}}"; }
    }

    /// <summary>
    /// Render the live session's current state to a complete anchor-stamped HTML document —
    /// the editor's full re-render (remount) without marshaling the saved bytes out to JS
    /// and back in. Same error convention as <see cref="RenderBlockHtml"/>: HTML starts
    /// with '&lt;', an error object starts with '{'.
    /// </summary>
    [JSExport]
    public static string RenderHtml(int h, string cssPrefix, bool fabricateClasses, bool paginated, double scale)
    {
        try { return DocxSessionOps.RenderHtml(h, cssPrefix, fabricateClasses, paginated, scale); }
        catch (System.Exception ex) { return $"{{\"error\":\"{JsonEncodedText.Encode(ex.Message ?? string.Empty)}\"}}"; }
    }

    [JSExport]
    public static string ReplaceText(int h, string anchor, string md) =>
        DocxSessionOps.ReplaceText(h, anchor, md);

    [JSExport]
    public static string DeleteBlock(int h, string anchor) =>
        DocxSessionOps.DeleteBlock(h, anchor);

    /// <summary>
    /// Bridge for <see cref="DocxSession.DeleteRange"/>. Deletes every top-level
    /// block-level sibling between <paramref name="fromAnchorId"/> (inclusive) and
    /// <paramref name="toAnchorIdExclusive"/> (exclusive). Both anchors must share a
    /// direct parent and live in the same package part. Returns a single EditResult.
    /// </summary>
    [JSExport]
    public static string DeleteRange(int h, string fromAnchorId, string toAnchorIdExclusive) =>
        DocxSessionOps.DeleteRange(h, fromAnchorId, toAnchorIdExclusive);

    /// <summary>
    /// Bridge for <see cref="DocxSession.DeleteSection"/>. Deletes a heading and
    /// every sibling below it up to (but not including) the next heading at the
    /// same or higher level. <paramref name="headingAnchorId"/> must address a
    /// heading-kind anchor (<c>h</c>).
    /// </summary>
    [JSExport]
    public static string DeleteSection(int h, string headingAnchorId) =>
        DocxSessionOps.DeleteSection(h, headingAnchorId);

    [JSExport]
    public static string InsertParagraph(int h, string anchor, string posStr, string md) =>
        DocxSessionOps.InsertParagraph(h, anchor, DocxSessionJson.ParsePos(posStr), md);

    [JSExport]
    public static string SplitParagraph(int h, string anchor, int offset) =>
        DocxSessionOps.SplitParagraph(h, anchor, offset);

    [JSExport]
    public static string MergeParagraphs(int h, string first, string second) =>
        DocxSessionOps.MergeParagraphs(h, first, second);

    /// <summary>
    /// Insert an empty paragraph carrying a bottom border (an S-1-style horizontal rule)
    /// before/after the anchor. <paramref name="ruleJson"/> is an optional border-edge object
    /// ({ style?, size?, color?, space? }); empty string uses the default rule.
    /// </summary>
    [JSExport]
    public static string InsertHorizontalRule(int h, string anchor, string posStr, string ruleJson) =>
        DocxSessionOps.InsertHorizontalRule(h, anchor, DocxSessionJson.ParsePos(posStr), ruleJson);

    /// <summary>
    /// Insert a rows×cols table before/after the anchor. <paramref name="optionsJson"/> is a
    /// TableInsertOptions object ({ borderless?, cellContents?: string[], cellAlignment? }).
    /// Returns an EditResult whose <c>created</c> lists the cell-paragraph anchors (row-major).
    /// </summary>
    [JSExport]
    public static string InsertTable(int h, string anchor, string posStr, int rows, int cols, string optionsJson) =>
        DocxSessionOps.InsertTable(h, anchor, DocxSessionJson.ParsePos(posStr), rows, cols, optionsJson);

    [JSExport]
    public static string InsertTableRow(int h, string cellAnchor, string posStr) =>
        DocxSessionOps.InsertTableRow(h, cellAnchor, DocxSessionJson.ParsePos(posStr));

    [JSExport]
    public static string InsertTableColumn(int h, string cellAnchor, string posStr) =>
        DocxSessionOps.InsertTableColumn(h, cellAnchor, DocxSessionJson.ParsePos(posStr));

    [JSExport]
    public static string DeleteTableRow(int h, string cellAnchor) =>
        DocxSessionOps.DeleteTableRow(h, cellAnchor);

    [JSExport]
    public static string DeleteTableColumn(int h, string cellAnchor) =>
        DocxSessionOps.DeleteTableColumn(h, cellAnchor);

    /// <summary>Retune the column widths of the table containing <paramref name="cellAnchor"/>.
    /// <paramref name="widthsJson"/> is a JSON array of per-column twip widths (one positive
    /// value per column, left→right).</summary>
    [JSExport]
    public static string SetColumnWidths(int h, string cellAnchor, string widthsJson) =>
        DocxSessionOps.SetColumnWidths(h, cellAnchor, widthsJson);

    /// <summary>Set the table-level borders of the table containing <paramref name="cellAnchor"/>.
    /// <paramref name="specJson"/> is a TableBorderSpec object
    /// ({ scope?: "all"|"outside"|"inside", style?, size?, color? }); "" = thin single all round.</summary>
    [JSExport]
    public static string SetTableBorders(int h, string cellAnchor, string specJson) =>
        DocxSessionOps.SetTableBorders(h, cellAnchor, specJson);

    /// <summary>Shade the cell containing <paramref name="cellAnchor"/> (scope "cell") or its whole
    /// row (scope "row"). <paramref name="fill"/> is a hex RRGGBB triplet or "auto"; "" clears.</summary>
    [JSExport]
    public static string SetCellShading(int h, string cellAnchor, string fill, string scope) =>
        DocxSessionOps.SetCellShading(h, cellAnchor, fill, scope);

    /// <summary>Mark/unmark the row containing <paramref name="cellAnchor"/> as a repeating
    /// header row (w:trPr/w:tblHeader).</summary>
    [JSExport]
    public static string SetRepeatHeaderRow(int h, string cellAnchor, bool repeat) =>
        DocxSessionOps.SetRepeatHeaderRow(h, cellAnchor, repeat);

    /// <summary>
    /// Set the section's running header story (<paramref name="anchor"/> = any body block in the
    /// section) to <paramref name="markdown"/>. <paramref name="kind"/> is "default" | "first" |
    /// "even". Creates the header part + reference if absent; returns the created header-paragraph
    /// anchors (scope <c>hdr{N}</c>) in <c>created</c>.
    /// </summary>
    [JSExport]
    public static string SetHeaderText(int h, string anchor, string kind, string markdown) =>
        DocxSessionOps.SetHeaderText(h, anchor, DocxSessionJson.ParseHeaderFooterKind(kind), markdown);

    /// <summary>Set the section's running footer story — see <see cref="SetHeaderText"/>; the created
    /// footer-paragraph anchors (scope <c>ftr{N}</c>) come back in <c>created</c>.</summary>
    [JSExport]
    public static string SetFooterText(int h, string anchor, string kind, string markdown) =>
        DocxSessionOps.SetFooterText(h, anchor, DocxSessionJson.ParseHeaderFooterKind(kind), markdown);

    /// <summary>Append a page-number field to the paragraph <paramref name="anchor"/> (typically a
    /// header/footer paragraph). <paramref name="field"/> is "currentPage" (PAGE) | "totalPages"
    /// (NUMPAGES). <paramref name="format"/> is an ST_NumberFormat token ("lowerRoman", …) writing
    /// the field's own <c>\*</c> switch, or "" for a plain field that follows the section's format
    /// (see <see cref="SetPageNumbering"/>) — which is the normal choice. Returns the affected
    /// paragraph anchor in <c>modified</c>.</summary>
    [JSExport]
    public static string InsertPageNumberField(int h, string anchor, string field, string format) =>
        DocxSessionOps.InsertPageNumberField(h, anchor,
            DocxSessionJson.ParsePageNumberField(field),
            DocxSessionJson.ParseNumberFormatOrNull(format));

    /// <summary>
    /// Bridge for <see cref="DocxSession.SetPageNumbering"/> — the section (<c>w:pgNumType</c>) half
    /// of page numbering, addressed by any body block in the section. <paramref name="opJson"/> is
    /// { start?: int, format?: ST_NumberFormat token }; omitted fields are left unchanged.
    /// </summary>
    [JSExport]
    public static string SetPageNumbering(int h, string anchor, string opJson) =>
        DocxSessionOps.SetPageNumbering(h, anchor, DocxSessionJson.ParsePageNumberingOp(opJson));

    /// <summary>Remove the section's page-numbering start/format — see
    /// <see cref="DocxSession.ClearPageNumbering"/>.</summary>
    [JSExport]
    public static string ClearPageNumbering(int h, string anchor) =>
        DocxSessionOps.ClearPageNumbering(h, anchor);

    /// <summary>Make the <paramref name="kind"/> ("first" | "even") header/footer stories of the
    /// section owning <paramref name="anchor"/> actually render — sets <c>w:titlePg</c> /
    /// <c>w:evenAndOddHeaders</c>. "default" is a no-op. Needed when a document already carries a
    /// first/even reference with the flag absent, which content writes alone don't fix.</summary>
    [JSExport]
    public static string EnsureHeaderFooterVisible(int h, string anchor, string kind) =>
        DocxSessionOps.EnsureHeaderFooterVisible(h, anchor, DocxSessionJson.ParseHeaderFooterKind(kind));

    /// <summary>
    /// Create a footnote with body <paramref name="markdown"/> and cite it from the body paragraph
    /// <paramref name="anchor"/> at <paramref name="characterOffset"/> characters into its text.
    /// Creates the footnotes part + Word's reserved separator notes on first use. Returns the created
    /// note anchors (kind <c>fn</c> plus its <c>p</c>/scope <c>fn</c> paragraphs) in <c>created</c>.
    /// </summary>
    [JSExport]
    public static string InsertFootnote(int h, string anchor, int characterOffset, string markdown) =>
        DocxSessionOps.InsertFootnote(h, anchor, characterOffset, markdown);

    /// <summary>Create an endnote — see <see cref="InsertFootnote"/>; writes the endnotes part and a
    /// <c>w:endnoteReference</c>, and the created definition anchor has kind <c>en</c>.</summary>
    [JSExport]
    public static string InsertEndnote(int h, string anchor, int characterOffset, string markdown) =>
        DocxSessionOps.InsertEndnote(h, anchor, characterOffset, markdown);

    /// <summary>
    /// Add a native Word comment on body paragraph <paramref name="anchor"/> (issue #300).
    /// <paramref name="spanJson"/> is <c>""</c> for the whole block or
    /// <c>{"start":int,"length":int}</c>; <paramref name="initials"/>/<paramref name="date"/>
    /// are <c>""</c> when absent (date is ISO-8601 and written only when provided). Returns the
    /// created definition anchor (kind <c>cmt</c>) plus its <c>p</c>/scope-<c>cmt</c> paragraph
    /// anchors in <c>created</c>.
    /// </summary>
    [JSExport]
    public static string AddComment(int h, string anchor, string spanJson, string author,
        string initials, string date, string markdown) =>
        DocxSessionOps.AddComment(h, anchor, ParseSpan(spanJson), author,
            string.IsNullOrEmpty(initials) ? null : initials,
            string.IsNullOrEmpty(date) ? null : date,
            markdown);

    /// <summary>Add a native reply with an adjacent reference; <c>w15:paraIdParent</c> links it
    /// to the immediate parent so it inherits the thread root's range.</summary>
    [JSExport]
    public static string AddCommentReply(int h, string parentCommentAnchor, string author,
        string initials, string date, string markdown) =>
        DocxSessionOps.AddCommentReply(h, parentCommentAnchor, author,
            string.IsNullOrEmpty(initials) ? null : initials,
            string.IsNullOrEmpty(date) ? null : date,
            markdown);

    /// <summary>Replace a comment's body text, addressed by its definition anchor (kind
    /// <c>cmt</c>); identity attributes (author/initials/date) are preserved.</summary>
    [JSExport]
    public static string UpdateComment(int h, string commentAnchor, string markdown) =>
        DocxSessionOps.UpdateComment(h, commentAnchor, markdown);

    /// <summary>Set <c>w15:done</c> for one comment (<c>false</c> reopens it), creating the
    /// paraId-keyed metadata parts when the comment was previously flat.</summary>
    [JSExport]
    public static string SetCommentResolved(int h, string commentAnchor, bool resolved) =>
        DocxSessionOps.SetCommentResolved(h, commentAnchor, resolved);

    /// <summary>Remove a comment: definition + body marker triple + threading entries.</summary>
    [JSExport]
    public static string RemoveComment(int h, string commentAnchor) =>
        DocxSessionOps.RemoveComment(h, commentAnchor);

    /// <summary>The document's comments in part order:
    /// <c>[{"anchorId","author","initials"?,"date"?,"text","parentAnchorId"?,"resolved"?}]</c>.</summary>
    [JSExport]
    public static string ListComments(int h) => DocxSessionOps.ListComments(h);

    /// <summary>Markup-native tracked-revision listing (issue #318), document order:
    /// <c>[{"id","type","author","date"?,"text","anchorId"?}]</c>. Ids are stable while
    /// the markup exists and address AcceptRevision/RejectRevision; type is
    /// <c>insert</c>/<c>delete</c>/<c>move</c>/<c>format</c>.</summary>
    [JSExport]
    public static string ListRevisions(int h) => DocxSessionOps.ListRevisions(h);

    /// <summary>Accept ONE revision by id (an undoable session mutation); returns an
    /// EditResult envelope.</summary>
    [JSExport]
    public static string AcceptRevision(int h, string revisionId) =>
        DocxSessionOps.AcceptRevision(h, revisionId);

    /// <summary>Reject ONE revision by id — the inverse of AcceptRevision.</summary>
    [JSExport]
    public static string RejectRevision(int h, string revisionId) =>
        DocxSessionOps.RejectRevision(h, revisionId);

    [JSExport]
    public static string ApplyFormat(int h, string anchor, string spanJson, string opJson) =>
        DocxSessionOps.ApplyFormat(h, anchor, ParseSpan(spanJson), DocxSessionJson.ParseFormatOp(opJson));

    /// <summary>
    /// Bridge for the substring-targeted <see cref="DocxSession.ApplyFormat(string, string, FormatOp)"/>
    /// overload. Lets JS callers say "bold the substring 'foo' in this paragraph" without
    /// computing offsets — the overload finds the first occurrence and converts to a CharSpan.
    /// </summary>
    [JSExport]
    public static string ApplyFormatBySubstring(int h, string anchor, string substring, string opJson) =>
        DocxSessionOps.ApplyFormatBySubstring(h, anchor, substring, DocxSessionJson.ParseFormatOp(opJson));

    [JSExport]
    public static string SetParagraphStyle(int h, string anchor, string styleId) =>
        DocxSessionOps.SetParagraphStyle(h, anchor, styleId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.SetParagraphFormat"/>. <paramref name="opJson"/> is
    /// { alignment?: "left"|"center"|"right"|"justify", indentDelta?: int (twips),
    /// firstLineIndent?/hangingIndent?: int (twips, mutually exclusive),
    /// spacingBefore?/spacingAfter?: int (twips), lineSpacing?: int,
    /// lineSpacingRule?: "auto"|"exact"|"atLeast", pageBreakBefore?: bool };
    /// omitted fields are left unchanged.
    /// </summary>
    [JSExport]
    public static string SetParagraphFormat(int h, string anchor, string opJson) =>
        DocxSessionOps.SetParagraphFormat(h, anchor, DocxSessionJson.ParseParagraphFormatOp(opJson));

    [JSExport]
    public static string SetListLevel(int h, string anchor, int delta) =>
        DocxSessionOps.SetListLevel(h, anchor, delta);

    [JSExport]
    public static string RemoveListMembership(int h, string anchor) =>
        DocxSessionOps.RemoveListMembership(h, anchor);

    /// <summary>
    /// Bridge for <see cref="DocxSession.ApplyListFormat"/>. Promotes a plain paragraph to a
    /// bullet/numbered list item (synthesizing a numbering definition if needed) or removes
    /// list membership. <paramref name="kind"/> is "bullet" | "decimal" | "lowerLetter" |
    /// "upperLetter" | "lowerRoman" | "upperRoman" | a "*Parenthesis" variant of the numbered
    /// formats (e.g. "decimalParenthesis" → "(1)") | "none".
    /// </summary>
    [JSExport]
    public static string ApplyListFormat(int h, string anchor, string kind) =>
        DocxSessionOps.ApplyListFormat(h, anchor, DocxSessionJson.ParseListFormat(kind));

    /// <summary>
    /// Bridge for <see cref="DocxSession.ApplyListFormatRange"/>. Applies one list format across
    /// the contiguous sibling run from <paramref name="firstAnchor"/> to
    /// <paramref name="lastAnchor"/> inclusive — every member shares ONE <c>w:num</c> instance so
    /// the numbering sequence stays intact. Same <paramref name="kind"/> tokens as
    /// <see cref="ApplyListFormat"/>.
    /// </summary>
    [JSExport]
    public static string ApplyListFormatRange(int h, string firstAnchor, string lastAnchor, string kind) =>
        DocxSessionOps.ApplyListFormatRange(h, firstAnchor, lastAnchor, DocxSessionJson.ParseListFormat(kind));

    /// <summary>
    /// Bridge for <see cref="DocxSession.SetListStartOverride"/>. Restarts the anchored list
    /// item's numbering at <paramref name="value"/> (Word's <em>Set Numbering Value…</em>) by
    /// writing a <c>w:startOverride</c> on a dedicated <c>w:num</c> and repointing the anchored
    /// item plus the following members of its sequence.
    /// </summary>
    [JSExport]
    public static string SetListStartOverride(int h, string anchor, int value) =>
        DocxSessionOps.SetListStartOverride(h, anchor, value);

    /// <summary>Bridge for <see cref="DocxSession.ClearListStartOverride"/> — removes the
    /// numbering restart from the anchored item's whole sequence.</summary>
    [JSExport]
    public static string ClearListStartOverride(int h, string anchor) =>
        DocxSessionOps.ClearListStartOverride(h, anchor);

    [JSExport]
    public static string ReplaceCellContent(int h, string anchor, string md) =>
        DocxSessionOps.ReplaceCellContent(h, anchor, md);

    [JSExport]
    public static string RawGetXml(int h, string anchor) => DocxSessionOps.RawGetXml(h, anchor);

    [JSExport]
    public static string RawInsertXml(int h, string anchor, string posStr, string xml) =>
        DocxSessionOps.RawInsertXml(h, anchor, DocxSessionJson.ParsePos(posStr), xml);

    [JSExport]
    public static string RawReplaceXml(int h, string anchor, string xml) =>
        DocxSessionOps.RawReplaceXml(h, anchor, xml);

    /// <summary>
    /// Bridge for <see cref="DocxSession.Grep"/>. <paramref name="optionsJson"/>
    /// accepts <c>{regexOptions?: number, scope?: number, contextChars?: number,
    /// whitespace?: number, boundary?: number}</c>; numeric values follow the .NET
    /// <see cref="System.Text.RegularExpressions.RegexOptions"/>, <see cref="ProjectionScopes"/>,
    /// <see cref="WhitespaceMode"/>, and <see cref="ContextBoundary"/> flag layouts.
    /// Missing fields use sensible defaults (no options, body-only, 80 chars of
    /// context, preserve whitespace, char-boundary).
    /// </summary>
    [JSExport]
    public static string Grep(int h, string pattern, string optionsJson)
    {
        ParseGrepOptions(optionsJson, out var regexOpts, out var scope, out var contextChars, out var whitespace, out var boundary);
        return DocxSessionOps.Grep(h, pattern, regexOpts, scope, contextChars, whitespace, boundary);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.GrepCrossBlock"/>. Same <paramref name="optionsJson"/>
    /// shape as <see cref="Grep"/>; returns a JSON array of CrossBlockMatch records (each
    /// carries <c>enclosingAnchors[]</c> + <c>slices[]</c>).
    /// </summary>
    [JSExport]
    public static string GrepCrossBlock(int h, string pattern, string optionsJson)
    {
        ParseGrepOptions(optionsJson, out var regexOpts, out var scope, out var contextChars, out var whitespace, out var boundary);
        return DocxSessionOps.GrepCrossBlock(h, pattern, regexOpts, scope, contextChars, whitespace, boundary);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceTextRange"/>. <paramref name="optionsJson"/>
    /// accepts <c>{ignoreCase?: boolean, maxReplacements?: number}</c>. Returns a
    /// JSON array of EditResult — one per attempted match.
    /// </summary>
    [JSExport]
    public static string ReplaceTextRange(int h, string anchor, string find, string replace, string optionsJson)
    {
        ReplaceOptions? opts = null;
        if (!string.IsNullOrEmpty(optionsJson))
        {
            using var doc = JsonDocument.Parse(optionsJson);
            var root = doc.RootElement;
            opts = new ReplaceOptions
            {
                IgnoreCase = DocxSessionJson.TryGetBool(root, "ignoreCase", false),
                MaxReplacements = root.TryGetProperty("maxReplacements", out var mr) && mr.ValueKind == JsonValueKind.Number
                    ? mr.GetInt32() : (int?)null,
            };
        }
        return DocxSessionOps.ReplaceTextRange(h, anchor, find, replace, opts);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceTextAtSpan"/> — the span-addressable
    /// variant that lets JS callers replace a specific Grep match (by its EnclosingAnchor
    /// id + Span coordinates) instead of every occurrence of its text.
    /// </summary>
    [JSExport]
    public static string ReplaceTextAtSpan(int h, string anchor, int spanStart, int spanLength, string replace) =>
        DocxSessionOps.ReplaceTextAtSpan(h, anchor, spanStart, spanLength, replace);

    /// <summary>
    /// Bridge for <see cref="DocxSession.ReplaceInner(TextMatch, string)"/>. Takes the
    /// match's text (so the shared core can locate the brackets) plus anchor+span (so
    /// it can dispatch to <see cref="DocxSession.ReplaceTextAtSpan"/>). Bracket
    /// parsing happens transport-side rather than serializing a full <see cref="TextMatch"/>
    /// — the existing wire shape already carries text + anchor + span via Grep results,
    /// so callers don't need anything they don't already have.
    /// </summary>
    [JSExport]
    public static string ReplaceInner(int h, string matchText, string anchor, int spanStart, int spanLength, string newInner) =>
        DocxSessionOps.ReplaceInner(h, matchText, anchor, spanStart, spanLength, newInner);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindPlaceholders"/>. <paramref name="kinds"/>
    /// uses the numeric layout of <see cref="PlaceholderKinds"/> (BlankFill=1,
    /// AlternativeClause=2, Instruction=4, All=7); 0 returns nothing. <paramref name="scope"/>
    /// uses the <see cref="ProjectionScopes"/> flag layout. Returns a JSON array of placeholders.
    /// </summary>
    [JSExport]
    public static string FindPlaceholders(int h, int kinds, int scope, int contextChars, int boundary) =>
        DocxSessionOps.FindPlaceholders(h, (PlaceholderKinds)kinds, (ProjectionScopes)scope, contextChars, (ContextBoundary)boundary);

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetEditSummary"/>. Returns a JSON object
    /// with placeholder, underscore-run, footnote, and comment counts useful for
    /// "am I done?" verification at the end of an edit pipeline.
    /// </summary>
    [JSExport]
    public static string GetEditSummary(int h) => DocxSessionOps.GetEditSummary(h);

    /// <summary>
    /// Bridge for <see cref="DocxSession.RemainingPlaceholders"/>. Discoverability
    /// alias for <see cref="FindPlaceholders"/> — same return shape.
    /// </summary>
    [JSExport]
    public static string RemainingPlaceholders(int h, int kinds) =>
        DocxSessionOps.RemainingPlaceholders(h, (PlaceholderKinds)kinds);

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetDiff"/>. <paramref name="format"/> uses
    /// the numeric layout of <see cref="DiffFormat"/>: <c>Json=0</c> (anchor-keyed JSON
    /// array), <c>Unified=1</c> (patch(1)-compatible text), <c>SideBySide=2</c> (two-column
    /// text). Unknown numeric values throw <c>NotSupportedException</c> on the .NET side,
    /// surfaced to JS as a thrown error.
    /// </summary>
    [JSExport]
    public static string GetDiff(int h, int format) =>
        DocxSessionOps.GetDiff(h, (DiffFormat)format);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByAnnotation"/>. Returns a JSON array of
    /// <see cref="AnchorTarget"/> records (each <c>{id, kind, scope, unid, partUri}</c>);
    /// empty array when the id is unknown.
    /// </summary>
    [JSExport]
    public static string FindByAnnotation(int h, string annotationId) =>
        DocxSessionOps.FindByAnnotation(h, annotationId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByLabel"/>. Returns a JSON object keyed by
    /// annotation id; each value is the same AnchorTarget array shape as
    /// <see cref="FindByAnnotation"/>.
    /// </summary>
    [JSExport]
    public static string FindByLabel(int h, string labelId) =>
        DocxSessionOps.FindByLabel(h, labelId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByBookmark"/>. Same return shape as
    /// <see cref="FindByAnnotation"/>; accepts any bookmark name (Docxodus-managed or
    /// user-authored).
    /// </summary>
    [JSExport]
    public static string FindByBookmark(int h, string bookmarkName) =>
        DocxSessionOps.FindByBookmark(h, bookmarkName);

    /// <summary>
    /// Bridge for <see cref="DocxSession.ListAnnotations"/>. Returns a JSON array of
    /// annotation records — id/labelId/label/color/author/created/bookmarkName/
    /// annotatedText, plus the metadata bag when non-empty. Page-info cache fields
    /// are omitted to keep the wire format compact; callers needing them can use the
    /// .NET API.
    /// </summary>
    [JSExport]
    public static string ListAnnotations(int h) => DocxSessionOps.ListAnnotations(h);

    // ─── Tier E: annotations (write surface) ──────────────────────────────

    /// <summary>
    /// Bridge for <see cref="DocxSession.AddAnnotation"/>. The span is encoded as
    /// a JSON string (empty/null = no span = annotate whole block, otherwise
    /// <c>{"start": int, "length": int}</c>) matching the existing
    /// <see cref="ApplyFormat"/> convention. The annotation JSON is a camelCase
    /// mirror of <see cref="DocumentAnnotation"/>; <see cref="DocxSessionJson.DeserializeAnnotation"/>
    /// parses it with <see cref="JsonDocument"/>, so this bridge is trim-safe under the
    /// WASM Release build.
    /// </summary>
    [JSExport]
    public static string AddAnnotation(int h, string anchorId, string spanJson, string annotationJson) =>
        DocxSessionOps.AddAnnotation(h, anchorId, ParseSpan(spanJson), annotationJson);

    /// <summary>
    /// Session-style RemoveAnnotation (distinct from the existing WmlDocument-style
    /// <see cref="RemoveAnnotation"/> which takes byte arrays). Removes the bookmark
    /// pair and custom-XML entry from the live session document.
    /// </summary>
    [JSExport]
    public static string SessionRemoveAnnotation(int h, string annotationId) =>
        DocxSessionOps.RemoveAnnotation(h, annotationId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.UpdateAnnotation"/>. Parsing is delegated to
    /// <see cref="DocxSessionJson.DeserializeAnnotationUpdate"/> (JsonDocument-based,
    /// trim-safe). <c>metadataPatch</c> honours explicit nulls — a null value removes
    /// the key, a missing key leaves it unchanged.
    /// </summary>
    [JSExport]
    public static string UpdateAnnotation(int h, string annotationId, string updateJson) =>
        DocxSessionOps.UpdateAnnotation(h, annotationId, updateJson);

    [JSExport]
    public static string MoveAnnotation(int h, string annotationId, string newAnchorId, string newSpanJson) =>
        DocxSessionOps.MoveAnnotation(h, annotationId, newAnchorId, ParseSpan(newSpanJson));

    /// <summary>Bridge for <see cref="DocxSession.Exists"/>. Returns true/false.</summary>
    [JSExport]
    public static bool Exists(int h, string anchorId) => DocxSessionOps.Exists(h, anchorId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetAnchorInfo"/>. Returns a JSON object
    /// <c>{id, kind, scope, textPreview}</c> or the literal <c>null</c> if the
    /// anchor is not found.
    /// </summary>
    [JSExport]
    public static string GetAnchorInfo(int h, string anchorId) => DocxSessionOps.GetAnchorInfo(h, anchorId);

    /// <summary>
    /// Bulk variant of <see cref="GetAnchorInfo"/>. Takes a JSON array of anchor
    /// ids and returns a JSON object keyed by id; each value is the AnchorInfo
    /// shape or <c>null</c> for unknown ids.
    /// </summary>
    [JSExport]
    public static string GetAnchorInfos(int h, string anchorIdsJson)
    {
        string[] ids;
        try
        {
            ids = JsonSerializer.Deserialize<string[]>(
                anchorIdsJson, DocxodusJsonContext.Default.StringArray) ?? System.Array.Empty<string>();
        }
        catch (JsonException)
        {
            return "{\"error\":\"malformed anchor id array\"}";
        }
        return DocxSessionOps.GetAnchorInfos(h, ids);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetBlockMetadata"/>. Returns a JSON
    /// object with style id/name, outline level, list membership (when present),
    /// and a hasInlineFormatting flag — or <c>"null"</c> if the anchor doesn't exist.
    /// </summary>
    [JSExport]
    public static string GetBlockMetadata(int h, string anchorId) =>
        DocxSessionOps.GetBlockMetadata(h, anchorId);

    /// <summary>
    /// Bulk variant of <see cref="GetBlockMetadata"/>. Takes a JSON array of anchor
    /// ids, returns a JSON object mapping each id to its metadata (or null).
    /// </summary>
    [JSExport]
    public static string GetBlockMetadatas(int h, string anchorIdsJson)
    {
        string[] ids;
        try
        {
            ids = JsonSerializer.Deserialize<string[]>(
                anchorIdsJson, DocxodusJsonContext.Default.StringArray) ?? System.Array.Empty<string>();
        }
        catch (JsonException)
        {
            return "{\"error\":\"malformed anchor id array\"}";
        }
        return DocxSessionOps.GetBlockMetadatas(h, ids);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetListMembership"/>. Returns a JSON
    /// object with numId/abstractNumId/level/format/etc., or <c>"null"</c>.
    /// </summary>
    [JSExport]
    public static string GetListMembership(int h, string anchorId) =>
        DocxSessionOps.GetListMembership(h, anchorId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetSectionInfo"/>. Returns a JSON object
    /// describing the governing <c>w:sectPr</c>, or <c>"null"</c> for non-body anchors.
    /// </summary>
    [JSExport]
    public static string GetSectionInfo(int h, string anchorId) =>
        DocxSessionOps.GetSectionInfo(h, anchorId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByText"/>. Returns a single AnchorTarget
    /// JSON object (first match in document order) or the literal <c>null</c> if no
    /// anchor contains the needle. <paramref name="optionsJson"/> accepts
    /// <c>{ignoreCase?, ignoreWhitespace?, kindFilter?, scopeFilter?}</c>.
    /// </summary>
    [JSExport]
    public static string FindByText(int h, string needle, string optionsJson) =>
        DocxSessionOps.FindByText(h, needle, ParseFindOptions(optionsJson));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindAllByText"/>. Same options shape as
    /// <see cref="FindByText"/>; returns the full AnchorTarget array in document order.
    /// </summary>
    [JSExport]
    public static string FindAllByText(int h, string needle, string optionsJson) =>
        DocxSessionOps.FindAllByText(h, needle, ParseFindOptions(optionsJson));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByRegex"/>. <paramref name="regexOptions"/>
    /// uses the numeric layout of <see cref="System.Text.RegularExpressions.RegexOptions"/>;
    /// <paramref name="optionsJson"/> matches the <see cref="FindByText"/> shape.
    /// </summary>
    [JSExport]
    public static string FindByRegex(int h, string pattern, int regexOptions, string optionsJson) =>
        DocxSessionOps.FindByRegex(h, pattern, (RegexOptions)regexOptions, ParseFindOptions(optionsJson));

    /// <summary>
    /// Bridge for <see cref="DocxSession.FindByKind"/>. <paramref name="scope"/> may be
    /// empty/null to match all scopes. No text scan — reads the AnchorIndex directly.
    /// </summary>
    [JSExport]
    public static string FindByKind(int h, string kind, string scope) =>
        DocxSessionOps.FindByKind(h, kind, string.IsNullOrEmpty(scope) ? null : scope);

    [JSExport]
    public static bool Undo(int h) => DocxSessionOps.Undo(h);

    [JSExport]
    public static bool Redo(int h) => DocxSessionOps.Redo(h);

    /// <summary>Switch how subsequent mutations are recorded (issue #304). mode is the
    /// numeric TrackedChangeMode (0=Accept, 1=RenderInline, 2=StripDeletions).</summary>
    [JSExport]
    public static void SetTrackedChanges(int h, int mode) =>
        DocxSessionOps.SetTrackedChanges(h, (TrackedChangeMode)mode);

    /// <summary>Change the author stamped on subsequent tracked-change markup.
    /// Empty string resets to the "docxodus" default (the AddComment null convention).</summary>
    [JSExport]
    public static void SetRevisionAuthor(int h, string author) =>
        DocxSessionOps.SetRevisionAuthor(h, string.IsNullOrEmpty(author) ? null : author);

    [JSExport]
    public static byte[] Save(int h) => DocxSessionOps.Save(h);

    /// <summary>Save KEEPING the projector's Unid bookkeeping — for the editor's remount, which
    /// re-renders these bytes and needs the anchors to survive the hop. ~6x larger than the
    /// document; never hand this to a user (see <see cref="Save"/>).</summary>
    [JSExport]
    public static byte[] SaveWithAnchorIds(int h) => DocxSessionOps.SaveWithAnchorIds(h);

    // ─── Helpers ────────────────────────────────────────────────────────

    private static CharSpan? ParseSpan(string json)
    {
        if (string.IsNullOrEmpty(json)) return null;
        using var doc = JsonDocument.Parse(json);
        return new CharSpan(
            doc.RootElement.GetProperty("start").GetInt32(),
            doc.RootElement.GetProperty("length").GetInt32());
    }

    private static FindOptions? ParseFindOptions(string optionsJson)
    {
        if (string.IsNullOrEmpty(optionsJson)) return null;
        using var doc = JsonDocument.Parse(optionsJson);
        return DocxSessionJson.ParseFindOptions(doc.RootElement);
    }

    private static void ParseGrepOptions(string optionsJson, out RegexOptions regexOpts,
        out ProjectionScopes scope, out int contextChars, out WhitespaceMode whitespace,
        out ContextBoundary boundary)
    {
        regexOpts = RegexOptions.None;
        scope = ProjectionScopes.Body;
        contextChars = 80;
        whitespace = WhitespaceMode.Preserve;
        boundary = ContextBoundary.Char;
        if (string.IsNullOrEmpty(optionsJson)) return;
        using var doc = JsonDocument.Parse(optionsJson);
        var root = doc.RootElement;
        if (root.TryGetProperty("regexOptions", out var ro) && ro.ValueKind == JsonValueKind.Number)
            regexOpts = (RegexOptions)ro.GetInt32();
        if (root.TryGetProperty("scope", out var s) && s.ValueKind == JsonValueKind.Number)
            scope = (ProjectionScopes)s.GetInt32();
        if (root.TryGetProperty("contextChars", out var c) && c.ValueKind == JsonValueKind.Number)
            contextChars = c.GetInt32();
        if (root.TryGetProperty("whitespace", out var w) && w.ValueKind == JsonValueKind.Number)
            whitespace = (WhitespaceMode)w.GetInt32();
        if (root.TryGetProperty("boundary", out var b) && b.ValueKind == JsonValueKind.Number)
            boundary = (ContextBoundary)b.GetInt32();
    }

}
