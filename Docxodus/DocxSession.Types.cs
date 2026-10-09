// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

// ─── Public value types ────────────────────────────────────────────────────

public enum Position { Before, After }

/// <summary>
/// How <see cref="DocxSession.Grep"/> and the <c>FindBy*</c> helpers treat Unicode
/// whitespace variants (NBSP, narrow NBSP, thin space) when matching. Word documents
/// routinely use NBSP between ordinals and colons (<c>First<NBSP>:</c>) so a needle
/// written with regular spaces silently misses without normalization — see issue #136.
/// </summary>
public enum WhitespaceMode
{
    /// <summary>Default: match against the document's original characters; NBSP stays NBSP.</summary>
    Preserve,

    /// <summary>Map U+00A0 / U+202F / U+2009 to ASCII space (U+0020) before matching.</summary>
    Normalize,
}

/// <summary>
/// Controls where <see cref="DocxSession.Grep"/> stops walking outward when
/// computing <see cref="TextMatch.ContextBefore"/> / <see cref="TextMatch.ContextAfter"/>.
/// The default <see cref="Char"/> just truncates at <c>contextChars</c>; the other
/// modes additionally stop at a natural-language boundary so the returned context
/// is unambiguously *this* match's surroundings, not text that belongs to an
/// adjacent placeholder or sibling sentence.
/// </summary>
public enum ContextBoundary
{
    /// <summary>No natural boundary; truncate at <c>contextChars</c> chars in each direction.
    /// Matches legacy behavior. This is the default.</summary>
    Char = 0,

    /// <summary>Stop at the nearest <c>'['</c> or <c>']'</c>. The dominant
    /// template-fill case: each placeholder's context is unambiguously its own,
    /// even when multiple placeholders crowd into one sentence.</summary>
    Bracket = 1,

    /// <summary>Stop at the nearest sentence-terminator (<c>. ! ? : ;</c>). Useful
    /// for callers building LLM prompts that want a self-contained snippet per match.</summary>
    Sentence = 2,

    /// <summary>Stop at the nearest comma. Useful for matches inside enumerations
    /// (<c>"X, Y, Z"</c>) where adjacent items are unambiguous siblings.</summary>
    Comma = 3,
}

public readonly record struct CharSpan(int Start, int Length);

/// <summary>The kind of target carried by a native Word hyperlink.</summary>
public enum HyperlinkKind { External, Internal }

/// <summary>A hyperlink destination. External targets are URI strings; internal targets are
/// bookmark names (without a leading <c>#</c>).</summary>
public sealed record HyperlinkTarget(HyperlinkKind Kind, string Target)
{
    public static HyperlinkTarget External(string uri) => new(HyperlinkKind.External, uri);
    public static HyperlinkTarget Internal(string bookmarkName) => new(HyperlinkKind.Internal, bookmarkName);
}

/// <summary>An unambiguous two-ended document range. Offsets are character boundaries and the
/// end is exclusive. Endpoints may be in different paragraphs but writable bookmark ranges must
/// remain in the same owning XML story part.</summary>
public sealed record DocumentRange(
    string StartAnchorId, int StartOffset, string EndAnchorId, int EndOffset)
{
    public static DocumentRange In(string anchorId, CharSpan span) =>
        new(anchorId, span.Start, anchorId, checked(span.Start + span.Length));
}

/// <summary>One paragraph-local slice of a bookmark's range.</summary>
public sealed record BookmarkRangeSegment(
    string OwningPartUri, string Scope, string AnchorId, CharSpan Span, string Text);

/// <summary>A first-class native Word hyperlink. <see cref="HyperlinkInfo.Id"/> follows the
/// session anchor identity contract: stable in-session and across <c>Save(true)</c> (or a session
/// configured with <c>PersistAnchorIds=true</c>), but not promised across the default stripped save.</summary>
public sealed record HyperlinkInfo(
    string Id, HyperlinkKind Kind, string OwningPartUri, string Scope,
    string AnchorId, CharSpan Span, string Text, string? Target,
    string? RelationshipId, bool? RelationshipIsExternal, bool IsBroken);

/// <summary>A native Word bookmark pair. <see cref="Range"/> identifies both endpoints even
/// when Word places them in different paragraphs; <see cref="Segments"/> provides the
/// paragraph-local text slices. Malformed or unmatched start markers are reported through
/// <see cref="BookmarkInfo.IsValid"/> and <see cref="BookmarkInfo.ValidationError"/>.</summary>
public sealed record BookmarkInfo(
    string Name, string BookmarkId, string StartPartUri, string StartScope,
    string? EndPartUri, string? EndScope, DocumentRange? Range,
    IReadOnlyList<BookmarkRangeSegment> Segments, string Text,
    bool IsPaired, bool IsManaged, bool IsValid, string? ValidationError);

public sealed record FormatOp
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    public bool? Strike { get; init; }
    public bool? Code { get; init; }
    public string? Color { get; init; }
    public string? RunStyle { get; init; }

    /// <summary>
    /// Vertical alignment (w:vertAlign): null = leave unchanged, "" / "none" / "baseline"
    /// = clear, "superscript" / "subscript" (or "super" / "sub") = set. Single-valued, so
    /// a string rather than a bool toggle.
    /// </summary>
    public string? VertAlign { get; init; }

    /// <summary>
    /// Font size in points (maps to <c>w:sz</c>/<c>w:szCs</c>, which store half-points).
    /// null = leave unchanged; a value &lt;= 0 clears the explicit size (falls back to the
    /// style/default). Fractional points are allowed (e.g. 7.5) and round to the nearest
    /// half-point. Needed for the S-1 cover page's large "FORM S-1" and company-name lines.
    /// </summary>
    public double? FontSizePts { get; init; }

    /// <summary>
    /// Run font family (maps to <c>w:rFonts</c> — sets <c>w:ascii</c>/<c>w:hAnsi</c>/<c>w:cs</c>
    /// to the name). null = leave unchanged; <c>""</c> clears the explicit font so the run
    /// inherits the style/default. Needed to match serif filings (e.g. an S-1 in Times New Roman).
    /// </summary>
    public string? FontFamily { get; init; }

    /// <summary>
    /// Text highlight colour (<c>w:highlight</c>) — one of Word's fixed <c>ST_HighlightColor</c>
    /// names: <c>yellow</c>, <c>green</c>, <c>cyan</c>, <c>magenta</c>, <c>blue</c>, <c>red</c>,
    /// <c>darkBlue</c>, <c>darkCyan</c>, <c>darkGreen</c>, <c>darkMagenta</c>, <c>darkRed</c>,
    /// <c>darkYellow</c>, <c>darkGray</c>, <c>lightGray</c>, <c>black</c>, <c>white</c>. null =
    /// leave unchanged; <c>""</c> or <c>"none"</c> clears. Any other value fails the op
    /// (<see cref="EditErrorCode.InternalError"/> carrying the message, the same way an invalid
    /// <see cref="VertAlign"/> does) — Word has no free-form highlight, only these swatches.
    /// </summary>
    public string? Highlight { get; init; }

    /// <summary>
    /// All-capitals display (<c>w:caps</c>). true sets it, false removes it, null leaves it. Word
    /// treats caps and small caps as one either/or slot, so setting this true also removes
    /// <c>w:smallCaps</c>.
    /// </summary>
    public bool? Caps { get; init; }

    /// <summary>
    /// Small-capitals display (<c>w:smallCaps</c>). true sets it, false removes it, null leaves
    /// it. Setting this true also removes <c>w:caps</c> (Word's either/or rule).
    /// </summary>
    public bool? SmallCaps { get; init; }
}

/// <summary>
/// Page geometry for <see cref="DocxSession.SetPageSetup"/> — the <c>w:pgSz</c> and
/// <c>w:pgMar</c> of the governing <c>w:sectPr</c>, which is what Word's <i>Page Setup</i> dialog
/// writes. Every field is tri-state: <c>null</c> leaves that attribute exactly as it is. All values
/// are twips (1440 = 1 inch).
/// </summary>
/// <remarks>
/// <see cref="Landscape"/> without an explicit size swaps the current width and height when they
/// are portrait-shaped (and back on <c>false</c>); with an explicit size it only writes
/// <c>w:orient</c>. Validation (<see cref="EditErrorCode.InvalidPageSetup"/>): page width and height
/// must be positive, margins and header/footer distances non-negative, and the opposing margins
/// must leave room for content (<c>left + right &lt; width</c>, <c>top + bottom &lt; height</c>),
/// measured against the section's effective values after the op.
/// </remarks>
public sealed record PageSetupOp
{
    /// <summary>Page width (<c>w:pgSz/@w:w</c>).</summary>
    public int? PageWidthTwips { get; init; }

    /// <summary>Page height (<c>w:pgSz/@w:h</c>).</summary>
    public int? PageHeightTwips { get; init; }

    /// <summary>
    /// Orientation (<c>w:pgSz/@w:orient</c>). <c>true</c> writes <c>landscape</c> and, when no
    /// explicit size is given and the page is currently taller than wide, swaps width and height;
    /// <c>false</c> removes the attribute and swaps back when the page is wider than tall.
    /// </summary>
    public bool? Landscape { get; init; }

    /// <summary>Top margin (<c>w:pgMar/@w:top</c>).</summary>
    public int? MarginTopTwips { get; init; }

    /// <summary>Bottom margin (<c>w:pgMar/@w:bottom</c>).</summary>
    public int? MarginBottomTwips { get; init; }

    /// <summary>Left margin (<c>w:pgMar/@w:left</c>).</summary>
    public int? MarginLeftTwips { get; init; }

    /// <summary>Right margin (<c>w:pgMar/@w:right</c>).</summary>
    public int? MarginRightTwips { get; init; }

    /// <summary>Header distance from the page's top edge (<c>w:pgMar/@w:header</c>).</summary>
    public int? HeaderDistanceTwips { get; init; }

    /// <summary>Footer distance from the page's bottom edge (<c>w:pgMar/@w:footer</c>).</summary>
    public int? FooterDistanceTwips { get; init; }
}

/// <summary>
/// One edge of a paragraph border (a <c>w:pBdr</c> child — <c>w:top</c>/<c>w:bottom</c>).
/// Drives the horizontal rules and section separators on an S-1-style cover page. When an
/// edge is set, null fields fall back to sensible defaults; use
/// <see cref="ParagraphFormatOp.ClearBorders"/> to remove all paragraph borders.
/// </summary>
public sealed record ParagraphBorderEdge
{
    /// <summary>Border line style (<c>w:val</c>): single, double, thick, dotted, dashed, … Default "single".</summary>
    public string? Style { get; init; }

    /// <summary>Border weight in eighths of a point (<c>w:sz</c>). Default 6 (≈0.75pt); a heavy rule ≈ 18–24.</summary>
    public int? Size { get; init; }

    /// <summary>Border color as a hex triplet without '#', or "auto" (<c>w:color</c>). Default "auto".</summary>
    public string? Color { get; init; }

    /// <summary>Padding between the border and the text in points (<c>w:space</c>). Default 1.</summary>
    public int? Space { get; init; }
}

/// <summary>Paragraph alignment (maps to w:jc): Justify → w:val "both".</summary>
public enum ParagraphAlignment { Left, Center, Right, Justify }

/// <summary>
/// How <see cref="ParagraphFormatOp.LineSpacing"/> is interpreted — maps to
/// <c>w:spacing/@w:lineRule</c>. Under <see cref="Auto"/> the value is in 240ths of a line
/// (240 = single, 360 = 1.5×, 480 = double); under <see cref="Exact"/>/<see cref="AtLeast"/>
/// it is a height in twips (20ths of a point — e.g. 480 = exactly 24pt).
/// </summary>
public enum LineSpacingRule { Auto, Exact, AtLeast }

/// <summary>
/// Which header/footer story a <see cref="DocxSession.SetHeaderText"/> /
/// <see cref="DocxSession.SetFooterText"/> call targets. Maps to the
/// <c>w:headerReference</c>/<c>w:footerReference</c> <c>w:type</c> attribute:
/// <list type="bullet">
///   <item><description><see cref="Default"/> — the story shown on every page that has no more specific override (<c>w:type="default"</c>).</description></item>
///   <item><description><see cref="First"/> — the first-page-only story (<c>w:type="first"</c>); the section's <c>w:titlePg</c> flag is set so Word honors it.</description></item>
///   <item><description><see cref="Even"/> — the even-page story (<c>w:type="even"</c>); <c>w:evenAndOddHeaders</c> is set in the settings part so Word honors it.</description></item>
/// </list>
/// Note that <c>w:evenAndOddHeaders</c> is document-global and governs footers too: once set,
/// even pages stop inheriting the Default footer, so a section with only a Default footer shows
/// no footer at all on even pages. Set an <see cref="Even"/> footer alongside an Even header if
/// footers should keep appearing on every page.
/// </summary>
public enum HeaderFooterKind { Default, First, Even }

/// <summary>
/// Which page-number field <see cref="DocxSession.InsertPageNumberField"/> emits:
/// <see cref="CurrentPage"/> → a <c>PAGE</c> field (the current page number),
/// <see cref="TotalPages"/> → a <c>NUMPAGES</c> field (the total page count),
/// <see cref="PageOfTotal"/> → Word's "Page X of Y" gallery entry: the text <c>Page </c>, a
/// <c>PAGE</c> field, the text <c> of </c> and a <c>NUMPAGES</c> field, every run inheriting the
/// paragraph's last run properties.
/// </summary>
public enum PageNumberField { CurrentPage, TotalPages, PageOfTotal }

/// <summary>
/// Section-level page-numbering setup for <see cref="DocxSession.SetPageNumbering"/> — the
/// <c>w:pgNumType</c> element, which is what Word's <i>Format Page Numbers…</i> dialog writes.
/// Each field is tri-state: <c>null</c> leaves that attribute exactly as it is (present or absent).
/// Use <see cref="DocxSession.ClearPageNumbering"/> to remove them.
/// </summary>
/// <remarks>
/// This governs how a <b>plain</b> <c>PAGE</c> field renders anywhere in the section, which is the
/// normal way to number pages: set the section once, insert unswitched fields. It is distinct from
/// the per-field <c>\*</c> switch <see cref="DocxSession.InsertPageNumberField"/> can stamp — that
/// overrides the section for one field, and a field carrying one stops following this setting.
/// </remarks>
public sealed record PageNumberingOp
{
    /// <summary>
    /// The page number this section starts at (<c>w:start</c>) — e.g. <c>1</c> to restart numbering
    /// at the section break, which is what front-matter/body splits need. <c>null</c> leaves the
    /// attribute unchanged; absent means the section continues the previous one's numbering.
    /// </summary>
    public int? Start { get; init; }

    /// <summary>
    /// The number format for this section's pages (<c>w:fmt</c>) — e.g.
    /// <see cref="NumberFormat.LowerRoman"/> for <c>i, ii, iii</c> front matter. <c>null</c> leaves
    /// the attribute unchanged; absent means Word's default (<c>1, 2, 3</c>).
    /// <see cref="NumberFormat.Bullet"/> is rejected — pages cannot be bulleted.
    /// </summary>
    public NumberFormat? Format { get; init; }
}

/// <summary>
/// Paragraph-level formatting for <see cref="DocxSession.SetParagraphFormat"/>. Each field
/// is tri-state: null leaves it unchanged. Alignment sets w:jc; PageBreakBefore toggles
/// w:pageBreakBefore (false removes); IndentDelta adjusts w:ind/@w:left by a twips delta
/// (clamped at 0), preserving any firstLine/hanging/right indents.
/// </summary>
public sealed record ParagraphFormatOp
{
    public ParagraphAlignment? Alignment { get; init; }
    public int? IndentDelta { get; init; }
    public bool? PageBreakBefore { get; init; }

    /// <summary>
    /// First-line indent in twips (<c>w:ind/@w:firstLine</c>; 1440 = 1 inch) — how far the
    /// paragraph's first line starts right of its left edge. Absolute, not a delta; 0 writes an
    /// explicit "no first-line indent" (overriding a style-inherited one). Word treats
    /// <c>w:firstLine</c>/<c>w:hanging</c> as one either/or slot, so setting this removes any
    /// <c>@w:hanging</c>, and an op setting both this and <see cref="HangingIndent"/> is rejected
    /// (<see cref="EditErrorCode.InvalidParagraphFormat"/>). Negative values are invalid
    /// (the attribute is unsigned in OOXML).
    /// </summary>
    public int? FirstLineIndent { get; init; }

    /// <summary>
    /// Hanging indent in twips (<c>w:ind/@w:hanging</c>; 1440 = 1 inch) — how far every line
    /// EXCEPT the first starts right of the paragraph's left edge. Mutually exclusive with
    /// <see cref="FirstLineIndent"/>; setting this removes any <c>@w:firstLine</c>. Absolute;
    /// 0 clears the hang explicitly; negatives are invalid.
    /// </summary>
    public int? HangingIndent { get; init; }

    /// <summary>Space above the paragraph in twips (<c>w:spacing/@w:before</c>; 20 twips = 1pt,
    /// so 240 = 12pt). Absolute; 0 writes an explicit zero; negatives are invalid.</summary>
    public int? SpacingBefore { get; init; }

    /// <summary>Space below the paragraph in twips (<c>w:spacing/@w:after</c>; 20 twips = 1pt).
    /// Absolute; 0 writes an explicit zero; negatives are invalid.</summary>
    public int? SpacingAfter { get; init; }

    /// <summary>
    /// Line spacing (<c>w:spacing/@w:line</c>). Units depend on <see cref="LineSpacingRule"/>:
    /// 240ths of a line under <see cref="Docxodus.LineSpacingRule.Auto"/> (240 = single,
    /// 360 = 1.5×, 480 = double), twips under <c>Exact</c>/<c>AtLeast</c>. Writes
    /// <c>@w:lineRule</c> alongside (defaulting to <c>auto</c> when
    /// <see cref="LineSpacingRule"/> is null); negatives are invalid.
    /// </summary>
    public int? LineSpacing { get; init; }

    /// <summary>How <see cref="LineSpacing"/> is interpreted (<c>w:spacing/@w:lineRule</c>).
    /// Only meaningful alongside <see cref="LineSpacing"/> — set without it, the op is rejected
    /// (<see cref="EditErrorCode.InvalidParagraphFormat"/>).</summary>
    public LineSpacingRule? LineSpacingRule { get; init; }

    /// <summary>Top paragraph border (<c>w:pBdr/w:top</c>). null = leave unchanged.</summary>
    public ParagraphBorderEdge? TopBorder { get; init; }

    /// <summary>Bottom paragraph border (<c>w:pBdr/w:bottom</c>). null = leave unchanged.
    /// This is what an S-1 horizontal rule is: an (often empty) paragraph with a bottom border.</summary>
    public ParagraphBorderEdge? BottomBorder { get; init; }

    /// <summary>When true, remove the entire <c>w:pBdr</c> (all paragraph borders) before applying
    /// any <see cref="TopBorder"/>/<see cref="BottomBorder"/> in this same op.</summary>
    public bool? ClearBorders { get; init; }
}

/// <summary>Options for <see cref="DocxSession.InsertTable"/>.</summary>
public sealed record TableInsertOptions
{
    /// <summary>When true, emit explicit "none" table + inside borders (an invisible layout table —
    /// the S-1 multi-column blocks). When false, a thin single border on every edge.</summary>
    public bool Borderless { get; init; }

    /// <summary>Row-major (row 0 left→right, then row 1, …) markdown for each cell. A null/short list
    /// leaves the remaining cells empty; each entry may parse to more than one paragraph.</summary>
    public IReadOnlyList<string>? CellContents { get; init; }

    /// <summary>Alignment applied to every cell paragraph (the S-1 columns are centered). null = leave default.</summary>
    public ParagraphAlignment? CellAlignment { get; init; }

    /// <summary>Per-column widths in twips (one per column, left→right). null = equal columns.
    /// A non-null list whose length != the column count is a caller error (rejected). Drives
    /// unequal layouts like the S-1's wide-left / narrow-right filing-header row.</summary>
    public IReadOnlyList<int>? ColumnWidths { get; init; }
}

/// <summary>
/// Options for <see cref="DocxSession.InsertTableOfContents"/>. Every field is a typed switch on
/// the underlying <c>TOC</c> field, so a caller never writes <c>\o "1-3"</c> by hand — a malformed
/// switch string renders as nothing in Word, silently, which is the failure this exists to prevent.
/// </summary>
public sealed record TableOfContentsOptions
{
    /// <summary>Heading levels to list (<c>\o</c>): a level or a range within 1-9, e.g. <c>"1-3"</c>
    /// for a three-level contract table of contents. A single level such as <c>"2"</c> is accepted
    /// and normalized to <c>"2-2"</c>.</summary>
    public string Levels { get; init; } = "1-3";

    /// <summary>Make each entry a hyperlink to its heading (<c>\h</c>). On by default: a table of
    /// contents nobody can click is a table of contents nobody uses.</summary>
    public bool Hyperlinks { get; init; } = true;

    /// <summary>Hide the leader tab and page numbers in Word's web view (<c>\z</c>), where page
    /// numbers are meaningless. On by default, matching what Word's own Insert &gt; Table of
    /// Contents writes.</summary>
    public bool HideTabAndPageNumbersInWeb { get; init; } = true;

    /// <summary>Include paragraphs carrying an outline level but no heading style (<c>\u</c>).
    /// On by default.</summary>
    public bool UseOutlineLevels { get; init; } = true;

    /// <summary>The heading above the table, in Word's <c>TOCHeading</c> style. <c>null</c> or empty
    /// inserts the table with no heading.</summary>
    public string? Title { get; init; } = "Contents";

    /// <summary>Position of the right-aligned, dot-leadered page-number tab stop, in twips.
    /// 9350 is Word's default for a US-Letter page with one-inch margins.</summary>
    public int RightTabPos { get; init; } = 9350;
}

/// <summary>Options for <see cref="DocxSession.InsertTableOfFigures"/>.</summary>
/// <remarks>A table of figures is a <c>TOC</c> field selecting by caption label rather than by
/// outline level — Word's own encoding, not a separate field type.</remarks>
public sealed record TableOfFiguresOptions
{
    /// <summary>The caption label whose captions the table lists (<c>\c</c>) — <c>"Figure"</c>,
    /// <c>"Table"</c>, <c>"Exhibit"</c>, or whatever the document's <c>SEQ</c> captions use.</summary>
    public string CaptionLabel { get; init; } = "Figure";

    /// <summary>Make each entry a hyperlink to its caption (<c>\h</c>).</summary>
    public bool Hyperlinks { get; init; } = true;

    /// <inheritdoc cref="TableOfContentsOptions.RightTabPos"/>
    public int RightTabPos { get; init; } = 9350;
}

/// <summary>Word's fixed table-of-authorities categories. The numbers are Word's, not ours: a TOA
/// field's <c>\c</c> switch selects one by its position.</summary>
public enum AuthorityCategory
{
    /// <summary>Cases — the default, and the one a brief's table of authorities usually leads with.</summary>
    Cases = 1,
    Statutes = 2,
    OtherAuthorities = 3,
    Rules = 4,
    Treatises = 5,
    Regulations = 6,
    ConstitutionalProvisions = 7,
}

/// <summary>Options for <see cref="DocxSession.InsertTableOfAuthorities"/>.</summary>
public sealed record TableOfAuthoritiesOptions
{
    /// <summary>Which category of authority the table lists (<c>\c</c>).</summary>
    public AuthorityCategory Category { get; init; } = AuthorityCategory.Cases;

    /// <summary>Make each entry a hyperlink to its citation (<c>\h</c>).</summary>
    public bool Hyperlinks { get; init; } = true;

    /// <summary>The separator between an entry and its page numbers (<c>\e</c>), e.g. <c>"\t"</c>
    /// for a tab. <c>null</c> leaves Word's default.</summary>
    public string? EntryPageSeparator { get; init; }

    /// <inheritdoc cref="TableOfContentsOptions.RightTabPos"/>
    public int RightTabPos { get; init; } = 9350;
}

/// <summary>Which table edges a <see cref="DocxSession.SetTableBorders"/> call targets:
/// <see cref="Outside"/> = top/left/bottom/right, <see cref="Inside"/> = the inner grid lines
/// (<c>w:insideH</c>/<c>w:insideV</c>), <see cref="All"/> = both.</summary>
public enum TableBorderScope { All, Outside, Inside }

/// <summary>Shading granularity for <see cref="DocxSession.SetCellShading"/>: the one cell the
/// anchor sits in, or every cell of its row (header-row banding).</summary>
public enum TableShadingScope { Cell, Row }

/// <summary>Height rule for <see cref="TableRowOptions.HeightTwips"/>. Values map to
/// <c>w:trHeight/@w:hRule</c>.</summary>
public enum TableRowHeightRule { Auto, AtLeast, Exact }

/// <summary>Row-level table properties written by <see cref="DocxSession.SetTableRowOptions"/>.
/// Null values leave the corresponding property untouched. A zero <see cref="HeightTwips"/>
/// removes any explicit row height.</summary>
public sealed record TableRowOptions
{
    public bool? RepeatHeader { get; init; }
    public bool? AllowBreakAcrossPages { get; init; }
    public int? HeightTwips { get; init; }
    public TableRowHeightRule HeightRule { get; init; } = TableRowHeightRule.AtLeast;
}

/// <summary>What <see cref="DocxSession.MergeCells"/> does with the content of the cells a merge
/// absorbs. <see cref="Append"/> (the default) moves their non-empty blocks into the surviving
/// cell — lossless; <see cref="Discard"/> drops them; <see cref="Reject"/> refuses the merge
/// (<see cref="EditErrorCode.InvalidTableMerge"/>) when any absorbed cell is non-empty.</summary>
public enum TableMergeContent { Append, Discard, Reject }

/// <summary>Options for <see cref="DocxSession.MergeCells"/>.</summary>
public sealed record TableMergeOptions
{
    /// <summary>How absorbed cells' content is handled. Default <see cref="TableMergeContent.Append"/>.</summary>
    public TableMergeContent Content { get; init; } = TableMergeContent.Append;
}

/// <summary>Border specification for <see cref="DocxSession.SetTableBorders"/>. Written as
/// explicit <c>w:tblBorders</c> edges, so it overrides any style-inherited borders; edges
/// outside <see cref="Scope"/> are left untouched.</summary>
public sealed record TableBorderSpec
{
    /// <summary>Which edges to write. Default <see cref="TableBorderScope.All"/>.</summary>
    public TableBorderScope Scope { get; init; } = TableBorderScope.All;

    /// <summary>Border line style (<c>w:val</c>): single, double, thick, dotted, dashed, … —
    /// or "none" to remove the targeted edges (written as explicit none, like
    /// <see cref="TableInsertOptions.Borderless"/>). Default "single".</summary>
    public string? Style { get; init; }

    /// <summary>Border weight in eighths of a point (<c>w:sz</c>). Default 4 (= 0.5pt), the same
    /// thin rule <see cref="DocxSession.InsertTable"/> writes.</summary>
    public int? Size { get; init; }

    /// <summary>Border color as a hex RRGGBB triplet without '#', or "auto" (<c>w:color</c>).
    /// Default "auto".</summary>
    public string? Color { get; init; }
}

/// <summary>
/// List membership for <see cref="DocxSession.ApplyListFormat"/> /
/// <see cref="DocxSession.ApplyListFormatRange"/>. The non-<see cref="None"/> members decompose
/// (via <c>Internal.NumberFormats.FromListFormat</c>) into an underlying <see cref="NumberFormat"/>
/// plus a parenthesized-level-text flag: <see cref="Decimal"/> renders <c>1.</c> while
/// <see cref="DecimalParenthesis"/> renders <c>(1)</c> — same <c>w:numFmt</c>, different
/// <c>w:lvlText</c>. The <c>*Parenthesis</c> variants are the legal-drafting presets
/// (<c>(a)</c>, <c>(i)</c>, <c>(1)</c>).
/// </summary>
public enum ListFormat
{
    None,
    Bullet,
    Decimal,
    LowerLetter,
    UpperLetter,
    LowerRoman,
    UpperRoman,
    DecimalParenthesis,
    LowerLetterParenthesis,
    UpperLetterParenthesis,
    LowerRomanParenthesis,
    UpperRomanParenthesis,
}

/// <summary>
/// Per-fragment visible formatting reported by <see cref="DocxSession.Grep"/>.
/// Booleans default to <c>false</c> meaning "not set on this fragment". The
/// fields cover what a callerlikely wants to preserve when rewriting a match in
/// place — character emphasis, color, hyperlink target, named run style.
/// </summary>
public sealed record RunFormatting
{
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    public bool Code { get; init; }
    public string? Color { get; init; }
    public string? HyperlinkUrl { get; init; }
    public string? RunStyle { get; init; }
}

/// <summary>
/// High-signal paragraph properties used by the formatting-inspection surface. Every property is
/// nullable so a <em>direct</em> snapshot can distinguish "not written here" from an explicit zero
/// or false. Effective snapshots fill the schema defaults for alignment, indentation, spacing,
/// line spacing, and on/off properties after applying document defaults and the paragraph-style
/// chain through <see cref="FormattingAssembler"/>.
/// </summary>
public sealed record ParagraphFormatting
{
    /// <summary>The paragraph style id in effect at this layer. This is accepted directly by
    /// <see cref="DocxSession.SetParagraphStyle"/>.</summary>
    public string? StyleId { get; init; }
    public ParagraphAlignment? Alignment { get; init; }
    public int? LeftIndentTwips { get; init; }
    public int? RightIndentTwips { get; init; }
    public int? FirstLineIndentTwips { get; init; }
    public int? HangingIndentTwips { get; init; }
    public int? SpacingBeforeTwips { get; init; }
    public int? SpacingAfterTwips { get; init; }
    public int? LineSpacing { get; init; }
    public LineSpacingRule? LineSpacingRule { get; init; }
    public bool? KeepNext { get; init; }
    public bool? KeepLines { get; init; }
    public bool? PageBreakBefore { get; init; }
    public int? OutlineLevel { get; init; }
    public string? ShadingFill { get; init; }
    public ParagraphBorderEdge? TopBorder { get; init; }
    public ParagraphBorderEdge? BottomBorder { get; init; }
}

/// <summary>
/// High-signal run properties used by style, anchor, and inline-span introspection. Nullable
/// fields preserve the difference between an absent direct property and an explicit off value;
/// effective snapshots resolve the document/style cascade and fill false for absent toggles.
/// </summary>
public sealed record RunFormattingInfo
{
    /// <summary>Character-style id at this layer. Accepted directly by
    /// <see cref="FormatOp.RunStyle"/>.</summary>
    public string? StyleId { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    public string? UnderlineStyle { get; init; }
    public bool? Strike { get; init; }
    public bool? Code { get; init; }
    public string? Color { get; init; }
    public string? Highlight { get; init; }
    public string? VertAlign { get; init; }
    public double? FontSizePts { get; init; }
    public string? FontFamily { get; init; }
    public bool? Caps { get; init; }
    public bool? SmallCaps { get; init; }
    public bool? Hidden { get; init; }
}

/// <summary>High-signal base properties for a table style. These describe the style definition,
/// not the geometry or formatting of any concrete table (owned by issue #450).</summary>
public sealed record TableStyleFormatting
{
    public string? Alignment { get; init; }
    public int? WidthTwips { get; init; }
    public int? IndentTwips { get; init; }
    public string? Layout { get; init; }
    public bool? HasBorders { get; init; }
    public string? CellShadingFill { get; init; }
}

/// <summary>One explicit style definition from the document's style catalog.</summary>
public sealed record StyleInfo
{
    /// <summary>Stable <c>w:styleId</c>, accepted by paragraph/run style mutation tools.</summary>
    required public string Id { get; init; }
    required public string Name { get; init; }

    /// <summary>OOXML style type (<c>paragraph</c>, <c>character</c>, <c>table</c>, or
    /// <c>numbering</c>).</summary>
    required public string Type { get; init; }
    public string? BasedOn { get; init; }
    public string? Next { get; init; }
    public bool IsDefault { get; init; }
    public bool IsCustom { get; init; }

    /// <summary>True when <c>w:latentStyles</c> has an exception for this style name. The following
    /// gallery fields resolve explicit style metadata over the exception and latent defaults.</summary>
    public bool HasLatentException { get; init; }
    public int? UiPriority { get; init; }
    public bool? SemiHidden { get; init; }
    public bool? UnhideWhenUsed { get; init; }
    public bool? QuickFormat { get; init; }
    public bool? Locked { get; init; }

    public ParagraphFormatting? ResolvedParagraph { get; init; }
    public RunFormattingInfo? ResolvedRun { get; init; }
    public TableStyleFormatting? ResolvedTable { get; init; }
}

/// <summary>
/// One text-bearing run inside a paragraph-like anchor. <see cref="AnchorId"/> plus
/// <see cref="Span"/> can be passed directly to <see cref="DocxSession.ApplyFormat"/>; the run
/// Unid is also reported for stable correlation but is not a separate mutation handle.
/// </summary>
public sealed record InlineSpan
{
    required public string AnchorId { get; init; }
    required public string RunUnid { get; init; }
    required public CharSpan Span { get; init; }
    required public string Text { get; init; }
    required public RunFormattingInfo Direct { get; init; }
    required public RunFormattingInfo Effective { get; init; }

    /// <summary>Outer-to-inner native content-control membership for this run. Empty when the
    /// run is not inside a w:sdt. Each id is directly accepted by content-control operations.</summary>
    public IReadOnlyList<string> ContentControlAnchorIds { get; init; } = Array.Empty<string>();
}

/// <summary>Direct and effective formatting for one paragraph-like anchor.</summary>
public sealed record FormattingInspection
{
    /// <summary>Stable anchor id accepted by paragraph and inline formatting mutation tools.</summary>
    required public string AnchorId { get; init; }
    required public ParagraphFormatting DirectParagraph { get; init; }
    required public ParagraphFormatting EffectiveParagraph { get; init; }
    required public IReadOnlyList<InlineSpan> Runs { get; init; }
}

/// <summary>
/// One piece of a <see cref="TextMatch"/> that came from a single <c>&lt;w:r&gt;</c> run.
/// The <see cref="Unid"/> uniquely identifies the run within the document; callers
/// rewriting the match can address each piece by its Unid + <see cref="SpanInElement"/>
/// and preserve the run's <see cref="Formatting"/> when constructing replacement XML.
/// </summary>
public sealed record RunFragment
{
    /// <summary>PtOpenXml.Unid of the <c>w:r</c> element this fragment came from.</summary>
    required public string Unid { get; init; }

    /// <summary>The text from this run that participates in the match.</summary>
    required public string Text { get; init; }

    /// <summary>Character offset + length of this fragment inside the run's flat text.</summary>
    required public CharSpan SpanInElement { get; init; }

    /// <summary>Visible formatting of the run this fragment came from.</summary>
    required public RunFormatting Formatting { get; init; }
}

/// <summary>
/// A single match returned by <see cref="DocxSession.Grep"/>. The match always lives
/// within one block-level element (the <see cref="EnclosingAnchor"/>); cross-block
/// matches aren't represented because OOXML doesn't allow text to span paragraphs.
/// </summary>
public sealed record TextMatch
{
    /// <summary>The matched text.</summary>
    required public string Text { get; init; }

    /// <summary>The smallest block-level anchor (paragraph/heading/list item/table cell) that fully contains the match.</summary>
    required public AnchorTarget EnclosingAnchor { get; init; }

    /// <summary>Character offset + length of the match in the enclosing block's flat text.</summary>
    required public CharSpan Span { get; init; }

    /// <summary>The run fragments the match spans, in document order. Always non-empty for a successful match.</summary>
    required public IReadOnlyList<RunFragment> Fragments { get; init; }

    /// <summary>Up to <c>contextChars</c> chars from the enclosing block immediately before the match.</summary>
    required public string ContextBefore { get; init; }

    /// <summary>Up to <c>contextChars</c> chars from the enclosing block immediately after the match.</summary>
    required public string ContextAfter { get; init; }

    /// <summary>Regex capture groups (index 0 is always the whole match; named groups appear at their numeric index).</summary>
    public IReadOnlyList<string> Groups { get; init; } = Array.Empty<string>();

    /// <summary>Null unless the search requested page citations.</summary>
    public PageCitation? Citation { get; init; }
}

/// <summary>
/// One block's contribution to a <see cref="CrossBlockMatch"/>. Each slice names the
/// block it came from, the offset+length of the matched substring within that block,
/// and the run-level fragment breakdown for that slice. A slice's <see cref="Fragments"/>
/// list is empty when the match touches an empty paragraph (e.g. the blank line between
/// two clauses) — the slice is still recorded so callers can see that the match
/// crossed the empty block.
/// </summary>
public sealed record BlockSlice
{
    /// <summary>The block-level anchor this slice belongs to.</summary>
    required public AnchorTarget Anchor { get; init; }

    /// <summary>Character offset + length of the slice within the block's own flat text.</summary>
    required public CharSpan SpanInBlock { get; init; }

    /// <summary>The run fragments contributing to this slice, in document order.</summary>
    required public IReadOnlyList<RunFragment> Fragments { get; init; }
}

/// <summary>
/// A single match returned by <see cref="DocxSession.GrepCrossBlock"/>. Unlike
/// <see cref="TextMatch"/>, the match may span multiple adjacent block-level elements
/// (paragraphs/headings/list items) under the same parent container. <see cref="Slices"/>
/// breaks the match down by block; <see cref="EnclosingAnchors"/> lists every block the
/// match touches, in document order.
/// </summary>
public sealed record CrossBlockMatch
{
    /// <summary>The matched text, including any block-boundary separators (<c>\n</c>) the regex matched across.</summary>
    required public string Text { get; init; }

    /// <summary>Every block-level anchor the match touches, in document order. Always non-empty.</summary>
    required public IReadOnlyList<AnchorTarget> EnclosingAnchors { get; init; }

    /// <summary>Per-block breakdown of the match, in document order. Always non-empty.</summary>
    required public IReadOnlyList<BlockSlice> Slices { get; init; }

    /// <summary>Up to <c>contextChars</c> chars from the surrounding concatenated text immediately before the match.</summary>
    required public string ContextBefore { get; init; }

    /// <summary>Up to <c>contextChars</c> chars from the surrounding concatenated text immediately after the match.</summary>
    required public string ContextAfter { get; init; }

    /// <summary>Regex capture groups (index 0 is always the whole match; named groups appear at their numeric index).</summary>
    public IReadOnlyList<string> Groups { get; init; } = Array.Empty<string>();

    /// <summary>One citation per <see cref="EnclosingAnchors"/> entry when requested.</summary>
    public IReadOnlyList<PageCitation>? Citations { get; init; }
}

/// <summary>Options that tune the <c>FindBy*</c> helpers on <see cref="DocxSession"/>.</summary>
public sealed record FindOptions
{
    /// <summary>Case-insensitive matching.</summary>
    public bool IgnoreCase { get; init; }

    /// <summary>Fold NBSP / narrow-NBSP / thin-space to ASCII space before matching (see <see cref="WhitespaceMode.Normalize"/>).</summary>
    public bool IgnoreWhitespace { get; init; }

    /// <summary>If set, only return anchors of this kind (e.g. <c>"h"</c> for headings).</summary>
    public string? KindFilter { get; init; }

    /// <summary>
    /// Coarse-grained scope filter — a flag set selecting whole categories of
    /// package parts (Body, all Headers, all Footers, Footnotes, Endnotes,
    /// Comments). Defaults to <see cref="ProjectionScopes.All"/>. Compose with
    /// <c>|</c> to widen, e.g. <c>Scopes = ProjectionScopes.Body | ProjectionScopes.Headers</c>.
    /// </summary>
    /// <remarks>Use this in preference to <see cref="ScopeFilter"/> — it's
    /// typed, composable, and uniform with <see cref="DocxSession.Grep"/>'s
    /// <c>scope</c> parameter. <see cref="ScopeFilter"/> remains for the rare
    /// case where you need to target a single named part like <c>"hdr1"</c>.</remarks>
    public ProjectionScopes Scopes { get; init; } = ProjectionScopes.All;

    /// <summary>If set, only return anchors whose scope name matches exactly
    /// (e.g. <c>"body"</c>, <c>"hdr1"</c>). Applied AFTER <see cref="Scopes"/>
    /// as a further narrowing — set both to restrict to one specific part inside
    /// a category. Most callers should use <see cref="Scopes"/> instead.</summary>
    public string? ScopeFilter { get; init; }

    /// <summary>Attach citations only if this exact registered layout is still valid.</summary>
    public PageCitationRequest? CitationRequest { get; init; }
}

/// <summary>Convenience predicates over the <see cref="ProjectionScopes"/> flag set.</summary>
public static class ProjectionScopesExtensions
{
    /// <summary>Returns true when <paramref name="scopeName"/> (e.g. <c>"body"</c>,
    /// <c>"hdr1"</c>, <c>"fn"</c>) belongs to <paramref name="set"/>.</summary>
    public static bool IncludesScope(this ProjectionScopes set, string scopeName)
    {
        if (set == ProjectionScopes.All) return true;
        if (string.IsNullOrEmpty(scopeName)) return false;
        if (scopeName == "body") return set.HasFlag(ProjectionScopes.Body);
        if (scopeName.StartsWith("hdr", System.StringComparison.Ordinal)) return set.HasFlag(ProjectionScopes.Headers);
        if (scopeName.StartsWith("ftr", System.StringComparison.Ordinal)) return set.HasFlag(ProjectionScopes.Footers);
        if (scopeName == "fn") return set.HasFlag(ProjectionScopes.Footnotes);
        if (scopeName == "en") return set.HasFlag(ProjectionScopes.Endnotes);
        if (scopeName == "cmt") return set.HasFlag(ProjectionScopes.Comments);
        return false;
    }
}

/// <summary>Options that tune <see cref="DocxSession.ReplaceTextRange"/>.</summary>
public sealed record ReplaceOptions
{
    /// <summary>Case-insensitive matching for the literal <c>find</c> needle.</summary>
    public bool IgnoreCase { get; init; }

    /// <summary>Cap the number of replacements; null = unlimited.</summary>
    public int? MaxReplacements { get; init; }

    /// <summary>Require exactly this many occurrences before applying any replacement.</summary>
    public int? ExpectedMatchCount { get; init; }

    /// <summary>Optional optimistic session/anchor guards evaluated before searching.</summary>
    public MutationPreconditions? Preconditions { get; init; }
}

/// <summary>
/// Options for <see cref="DocxSession.FillPlaceholders"/>.
/// </summary>
public sealed record FillOptions
{
    /// <summary>Which placeholder kinds to fill. Defaults to
    /// <see cref="PlaceholderKinds.All"/> so the picker is invoked for every kind
    /// the doc contains — <c>BlankFill</c>, <c>Instruction</c>, *and*
    /// <c>AlternativeClause</c>. Narrow with e.g. <c>BlankFill | Instruction</c>
    /// if you only want value-slot fills and intend to ignore bracketed clauses.</summary>
    /// <remarks>The previous default (<c>BlankFill | Instruction</c>) silently
    /// excluded <c>AlternativeClause</c> placeholders, which caused pickers with
    /// bracket-stripping rules to appear to do nothing on those matches. The new
    /// default lets the picker see everything; pickers that don't recognize a
    /// kind should simply return <c>null</c> for it.</remarks>
    public PlaceholderKinds Kinds { get; init; } = PlaceholderKinds.All;

    /// <summary>Which package parts to scan. Defaults to body.</summary>
    public ProjectionScopes Scope { get; init; } = ProjectionScopes.Body;

    /// <summary>Maximum iteration passes. <see cref="DocxSession.FindPlaceholders"/> returns
    /// innermost brackets only; stripping one layer can surface a previously-nested
    /// outer layer, so multi-pass iteration is sometimes needed. The default of 8
    /// is a safety cap against infinite loops on adversarial input. Set higher if
    /// you have deeply-nested templates.</summary>
    public int MaxPasses { get; init; } = 8;

    /// <summary>When <c>true</c> (default), if the placeholder match text starts
    /// with <c>"$"</c> (the regex <c>\$?\[…\]</c> captured a leading dollar sign)
    /// and the picker's return value does not start with <c>"$"</c>, the dollar
    /// is preserved by prepending it to the replacement. Set to <c>false</c> if
    /// you want full control over the replacement and to overwrite the <c>$</c>.</summary>
    public bool PreserveDollarPrefix { get; init; } = true;

    /// <summary>Threaded through to <see cref="DocxSession.FindPlaceholders"/> calls
    /// inside the multi-pass loop. Default 80 (matches the new Grep default).</summary>
    public int ContextChars { get; init; } = 80;

    /// <summary>Boundary mode for the per-match context windows the picker sees.
    /// Default <see cref="ContextBoundary.Char"/> (legacy truncate-at-contextChars).
    /// Pickers that rely on bracket-bounded context can opt into
    /// <see cref="ContextBoundary.Bracket"/> for unambiguous per-placeholder context.</summary>
    public ContextBoundary Boundary { get; init; } = ContextBoundary.Char;

    /// <summary>When the picker returns an empty string — the canonical "drop
    /// this optional clause entirely" signal — the placeholder span is deleted
    /// verbatim, which leaves whitespace and punctuation around the (now-gone)
    /// brackets untouched. The repro from issue #188:
    /// <c>"… on [date] [under the name [name]]."</c> with the outer wrapper
    /// dropped (picker returns <c>""</c>) becomes <c>"… on March 14, 2024 ."</c>
    /// — note the stray space before the period.
    /// <para>
    /// When this flag is <c>true</c>, an empty fill additionally absorbs adjacent
    /// chars based on the immediate neighbors of the placeholder span in the
    /// enclosing block's flat text:
    /// </para>
    /// <list type="bullet">
    ///   <item>Whitespace on both sides → consume the trailing space, so
    ///   <c>"alpha [opt] beta"</c> becomes <c>"alpha beta"</c> (one space) rather
    ///   than <c>"alpha  beta"</c> (two).</item>
    ///   <item>Whitespace before + clause-terminating punctuation
    ///   (<c>. , ; : ! ?</c>) after → drop the leading space, so
    ///   <c>"… 2024 [opt]."</c> becomes <c>"… 2024."</c>.</item>
    ///   <item>Open-bracket (<c>( [ {</c>) before + matching close-bracket
    ///   (<c>) ] }</c>) after → drop both, so an outer wrapper around a now-empty
    ///   inner (<c>"[[opt]]"</c>) doesn't leave bare brackets.</item>
    /// </list>
    /// Default <c>false</c> (preserve the legacy literal-delete behavior).
    /// $-prefix preservation (<see cref="PreserveDollarPrefix"/>) runs first,
    /// so a picker returning <c>""</c> for <c>$[xxx]</c> with the default
    /// <see cref="PreserveDollarPrefix"/> = <c>true</c> ends up replacing with
    /// <c>"$"</c> (not empty) and coalescing is skipped — that's intentional;
    /// set <see cref="PreserveDollarPrefix"/> = <c>false</c> when you want
    /// the <c>$</c> to drop along with the brackets.
    /// </summary>
    public bool CoalesceWhitespaceAroundEmptyFill { get; init; }
}

/// <summary>
/// Aggregate result envelope returned by <see cref="DocxSession.FillPlaceholders"/>.
/// </summary>
public sealed record BulkEditResult
{
    /// <summary>Number of placeholders filled by the picker.</summary>
    public int Filled { get; init; }

    /// <summary>Number of placeholders for which the picker returned <c>null</c>
    /// (counted once per placeholder, in the first pass that saw it). This is
    /// <em>not</em> a trustworthy "did the fill leave anything undone?" signal —
    /// a placeholder the picker said <c>null</c> to in pass 1 may be fully
    /// resolved by pass 2 (e.g. a nested-outer wrapper becomes fillable once
    /// its inner is stripped, or a structural delete removes the placeholder
    /// entirely). Use <see cref="StillPresent"/> for the "is the template
    /// done?" check, and consult <see cref="Unfilled"/> for the per-placeholder
    /// detail.</summary>
    public int Skipped { get; init; }

    /// <summary>Number of placeholders matching <see cref="FillOptions.Kinds"/>
    /// in <see cref="FillOptions.Scope"/> that remain in the document after the
    /// final pass. This is the metric to assert on when you want to know
    /// whether the template is fully filled — <c>0</c> means every placeholder
    /// the loop visited is now gone (filled, stripped, or removed by a
    /// structural edit). Unlike <see cref="Skipped"/>, this is taken from the
    /// post-loop document state, so multi-pass convergence is reflected
    /// correctly: <c>Skipped &gt; 0</c> together with <c>StillPresent = 0</c> means
    /// "picker said no the first time but later passes finished the job."
    /// Computed via a single <see cref="DocxSession.FindPlaceholders"/> call
    /// scoped to the same kinds/scope the loop was operating on.</summary>
    public int StillPresent { get; init; }

    /// <summary>The highest iteration pass that actually filled at least one
    /// placeholder matching <see cref="FillOptions.Kinds"/>. <c>1</c> means a
    /// single pass did all the work; higher values mean multi-pass nested-bracket
    /// stripping or partial picker convergence. <c>0</c> means no fills happened
    /// — either no placeholders matched at all (the scope/kinds filter returned
    /// nothing on the first scan) or every match's picker call returned <c>null</c>.</summary>
    public int Passes { get; init; }

    /// <summary>Placeholders the picker returned <c>null</c> for.</summary>
    public IReadOnlyList<TemplatePlaceholder> Unfilled { get; init; } = Array.Empty<TemplatePlaceholder>();

    /// <summary>Per-replacement failures. Populated when <see cref="DocxSession.ReplaceMatch"/>
    /// returned <c>Success = false</c> for an attempted fill.</summary>
    public IReadOnlyList<EditError> Errors { get; init; } = Array.Empty<EditError>();
}

/// <summary>
/// Categories of bracketed placeholders that <see cref="DocxSession.FindPlaceholders"/>
/// recognizes. Templates routinely mix these — a real-world COI has dozens of value
/// blanks, dozens of optional clauses, and dozens of drafter hints, all inside
/// square brackets — and an agent fills each kind differently.
/// </summary>
public enum PlaceholderKind
{
    /// <summary><c>[_______]</c> or <c>$[_______]</c> — a value slot the agent fills with text.</summary>
    BlankFill,

    /// <summary><c>[entire clause text in brackets]</c> — an optional clause the agent keeps or strips.</summary>
    AlternativeClause,

    /// <summary><c>[insert X]</c>, <c>[specify Y]</c>, <c>[*italicized hint*]</c> — a drafter hint the agent treats as a parameter description.</summary>
    Instruction,
}

/// <summary>Flag set for narrowing <see cref="DocxSession.FindPlaceholders"/>.</summary>
[System.Flags]
public enum PlaceholderKinds
{
    BlankFill = 1,
    AlternativeClause = 2,
    Instruction = 4,
    All = BlankFill | AlternativeClause | Instruction,
}

/// <summary>
/// A single placeholder found by <see cref="DocxSession.FindPlaceholders"/>. Wraps the
/// underlying <see cref="TextMatch"/> with a classified <see cref="Kind"/> and (for
/// <see cref="PlaceholderKind.Instruction"/> placeholders) a parsed <see cref="Hint"/>.
/// </summary>
public sealed record TemplatePlaceholder
{
    required public TextMatch Match { get; init; }
    required public PlaceholderKind Kind { get; init; }

    /// <summary>For <see cref="PlaceholderKind.Instruction"/>: the inner text with
    /// surrounding brackets/asterisks stripped (e.g. <c>"[insert percentage]"</c> →
    /// <c>"insert percentage"</c>; <c>"[*specify name*]"</c> → <c>"specify name"</c>).
    /// <c>null</c> for other kinds.</summary>
    public string? Hint { get; init; }

    /// <summary>
    /// Additional plausible classifications when the primary <see cref="Kind"/> is
    /// borderline. Empty by default; populated when a secondary heuristic also
    /// matches the placeholder text. The classic case is a long bracketed clause
    /// that happens to contain a <c>_______</c> blank: primary <see cref="Kind"/>
    /// is <see cref="PlaceholderKind.BlankFill"/> for back-compat, with
    /// <see cref="PlaceholderKind.AlternativeClause"/> in <c>AlternativeKinds</c>
    /// so callers can detect the ambiguity and treat the placeholder as a clause
    /// (strip brackets, then fill the inner blank).
    /// </summary>
    public IReadOnlyList<PlaceholderKind> AlternativeKinds { get; init; } = Array.Empty<PlaceholderKind>();
}

public sealed record AnchorInfo(string Id, string Kind, string Scope, string TextPreview)
{
    /// <summary>
    /// Resolved auto-numbering prefix (e.g. <c>"First"</c>, <c>"1."</c>). <c>null</c>
    /// when the element has no numbering or the kind doesn't carry it. See
    /// <see cref="AnchorTarget.AutoNumberPrefix"/> for the full rationale.
    /// </summary>
    public string? AutoNumberPrefix { get; init; }

    /// <summary>Hash of the live anchor subtree, excluding projector Unids and note ids.</summary>
    public string ContentHash { get; init; } = string.Empty;

    /// <summary>Exact visible text used by optimistic preconditions (never preview-truncated).</summary>
    public string VisibleText { get; init; } = string.Empty;

    /// <summary>What a reader sees: <see cref="AutoNumberPrefix"/> + space + <see cref="TextPreview"/>
    /// when a prefix is present, otherwise just <see cref="TextPreview"/>.</summary>
    public string FullText =>
        string.IsNullOrEmpty(AutoNumberPrefix)
            ? TextPreview
            : string.IsNullOrEmpty(TextPreview)
                ? AutoNumberPrefix!
                : AutoNumberPrefix + " " + TextPreview;
}

/// <summary>
/// The six list formats supported by the list write surface
/// (<c>InsertNumberedList</c>, <c>ConvertToNumberedList</c>, …) and
/// surfaced on <see cref="ListMembership.Format"/>. Maps to OOXML
/// <c>w:numFmt</c> values: <c>Decimal</c> → <c>decimal</c>,
/// <c>UpperLetter</c> → <c>upperLetter</c>, <c>LowerLetter</c> →
/// <c>lowerLetter</c>, <c>UpperRoman</c> → <c>upperRoman</c>,
/// <c>LowerRoman</c> → <c>lowerRoman</c>, <c>Bullet</c> → <c>bullet</c>.
/// Other OOXML formats resolve to <c>Decimal</c> (the safest fallback).
/// </summary>
public enum NumberFormat
{
    Decimal,
    UpperLetter,
    LowerLetter,
    UpperRoman,
    LowerRoman,
    Bullet,
}

/// <summary>
/// Numbering facts for a list-item paragraph. Returned by
/// <see cref="DocxSession.GetListMembership"/> and surfaced as
/// <see cref="BlockMetadata.List"/>.
/// </summary>
public sealed record ListMembership
{
    /// <summary>The stable paragraph/list-item anchor accepted by list mutation tools.</summary>
    required public string AnchorId { get; init; }

    /// <summary>The <c>w:numId</c> the paragraph belongs to (the <c>w:num</c> instance).</summary>
    required public int NumId { get; init; }

    /// <summary>The <c>w:abstractNumId</c> the paragraph's <c>w:num</c> points at (the format template).</summary>
    required public int AbstractNumId { get; init; }

    /// <summary>The paragraph's level (<c>w:ilvl</c>), 0-8.</summary>
    required public int Level { get; init; }

    /// <summary>The resolved <see cref="NumberFormat"/> for this paragraph's level.</summary>
    required public NumberFormat Format { get; init; }

    /// <summary>The start-override applied to this paragraph's level via
    /// <c>w:lvlOverride/w:startOverride</c>, if any. <c>null</c> when no override is in effect.</summary>
    public int? StartOverride { get; init; }

    /// <summary>The level definition's <c>w:start</c> value. Defaults to 1 when omitted.</summary>
    required public int Start { get; init; }

    /// <summary>The level's marker template (<c>w:lvlText</c>), e.g. <c>"%1."</c> or
    /// <c>"(%2)"</c>.</summary>
    public string? LevelText { get; init; }

    /// <summary>Numbering-level indentation from <c>w:lvl/w:pPr/w:ind</c>.</summary>
    public int? LeftIndentTwips { get; init; }
    public int? RightIndentTwips { get; init; }
    public int? FirstLineIndentTwips { get; init; }
    public int? HangingIndentTwips { get; init; }

    /// <summary>Always <c>true</c> for a paragraph carrying <c>w:numPr</c> (inline or via style).</summary>
    required public bool IsAutoNumbered { get; init; }

    /// <summary><c>true</c> when the <c>w:numPr</c> is inherited from the paragraph style chain
    /// rather than set directly on the paragraph. <c>false</c> when set inline on the paragraph.</summary>
    required public bool FromStyle { get; init; }

    /// <summary>The rendered auto-number prefix (e.g. <c>"1."</c>, <c>"(a)"</c>) — same value
    /// surfaced as <see cref="AnchorInfo.AutoNumberPrefix"/>. Duplicated here so callers don't
    /// have to take two round-trips.</summary>
    public string? GeneratedLabel { get; init; }
}

/// <summary>
/// Block-level structural metadata. Returned by <see cref="DocxSession.GetBlockMetadata"/>.
/// </summary>
public sealed record BlockMetadata
{
    /// <summary>Same as <see cref="AnchorInfo.Id"/> — the markdown-projection anchor id.</summary>
    required public string AnchorId { get; init; }

    /// <summary>Same as <see cref="AnchorInfo.Kind"/> — e.g. <c>"p"</c>, <c>"h"</c>, <c>"li"</c>, <c>"tc"</c>, <c>"tbl"</c>.</summary>
    required public string Kind { get; init; }

    /// <summary>Same as <see cref="AnchorInfo.Scope"/> — e.g. <c>"body"</c>, <c>"hdr1"</c>, <c>"fn"</c>.</summary>
    required public string Scope { get; init; }

    /// <summary>The <c>w:pStyle/@w:val</c> for paragraph kinds, or <c>w:tblStyle</c> for tables.
    /// <c>null</c> when no style is applied.</summary>
    public string? StyleId { get; init; }

    /// <summary>Resolved <c>w:name/@w:val</c> for <see cref="StyleId"/> from the styles part.
    /// <c>null</c> when styles part is absent or the style isn't defined.</summary>
    public string? StyleName { get; init; }

    /// <summary>Outline level: <c>w:pPr/w:outlineLvl</c> when present; otherwise
    /// inferred from a Heading1..Heading9 style (level 0..8); <c>null</c> otherwise.
    /// Word's outlineLvl is 0-based (0 = top heading).</summary>
    public int? OutlineLevel { get; init; }

    /// <summary>Populated for list-item paragraphs; <c>null</c> otherwise.</summary>
    public ListMembership? List { get; init; }

    /// <summary><c>true</c> when any descendant <c>w:r</c> carries a non-empty <c>w:rPr</c>
    /// (bold, italic, color, run style, etc.). Coarse but useful as a "does this paragraph
    /// have inline formatting at all?" probe.</summary>
    required public bool HasInlineFormatting { get; init; }
}

/// <summary>
/// A <c>w:headerReference</c>/<c>w:footerReference</c> on a section: which story kind it
/// supplies and the URI of the part holding that story. Lets a caller map a
/// <see cref="HeaderFooterKind"/> to a part — and thence to that part's projection anchors,
/// which carry the same <c>PartUri</c> — instead of guessing from part-collection order,
/// which carries no kind information.
/// </summary>
public sealed record HeaderFooterRef
{
    /// <summary>The reference's <c>w:type</c>. The attribute is optional in OOXML; an absent
    /// (or unrecognized) value means <see cref="HeaderFooterKind.Default"/>.</summary>
    required public HeaderFooterKind Kind { get; init; }

    /// <summary>URI of the header/footer part this reference points at.</summary>
    required public string PartUri { get; init; }

    /// <summary>
    /// <c>true</c> when this section declares no reference of <see cref="Kind"/> itself and the
    /// story is INHERITED from the nearest preceding section that does (ECMA-376 §17.6.17 — a
    /// section without a reference of a type continues the previous section's). Editing an
    /// inherited story edits the part both sections share, which is what Word does.
    /// </summary>
    public bool Inherited { get; init; }
}

/// <summary>
/// Page-layout snapshot for the <c>w:sectPr</c> that governs an anchor.
/// Returned by <see cref="DocxSession.GetSectionInfo"/>; <c>null</c> for
/// anchors outside the body part (footnotes/endnotes/headers/footers/comments).
/// </summary>
public sealed record SectionInfo
{
    /// <summary>The body anchor used for this query. It is stable within the session and accepted
    /// directly by section mutation tools such as <see cref="DocxSession.SetPageNumbering"/> and
    /// <see cref="DocxSession.SetHeaderText"/>.</summary>
    required public string AnchorId { get; init; }

    /// <summary>The Unid of the <c>w:sectPr</c> element this info describes. Stable across mutations.</summary>
    required public string SectionUnid { get; init; }

    required public int PageWidthTwips { get; init; }
    required public int PageHeightTwips { get; init; }
    required public bool Landscape { get; init; }
    required public int MarginTopTwips { get; init; }
    required public int MarginBottomTwips { get; init; }
    required public int MarginLeftTwips { get; init; }
    required public int MarginRightTwips { get; init; }

    /// <summary>Word's header/footer distance (0.5") when <c>w:pgMar</c> omits the attribute —
    /// the one declaration every reader and writer of the two distances falls back to.</summary>
    public const int DefaultHeaderFooterDistanceTwips = 720;

    /// <summary>Header distance from the page's top edge (<c>w:pgMar/@w:header</c>);
    /// <see cref="DefaultHeaderFooterDistanceTwips"/> when the attribute is absent.</summary>
    required public int HeaderDistanceTwips { get; init; }

    /// <summary>Footer distance from the page's bottom edge (<c>w:pgMar/@w:footer</c>);
    /// <see cref="DefaultHeaderFooterDistanceTwips"/> when the attribute is absent.</summary>
    required public int FooterDistanceTwips { get; init; }

    /// <summary>
    /// Word's "Different first page" flag — <c>true</c> when the governing <c>w:sectPr</c> carries
    /// an on-valued <c>w:titlePg</c>, i.e. the section's first-page header/footer stories actually
    /// render. Toggle with <see cref="DocxSession.SetHeaderFooterKindEnabled"/>.
    /// </summary>
    required public bool TitlePage { get; init; }

    /// <summary>
    /// Word's "Different odd &amp; even pages" flag — <c>true</c> when the settings part carries an
    /// on-valued <c>w:evenAndOddHeaders</c>. Document-global (it is a settings-part element), so
    /// every section reports the same value. Toggle with
    /// <see cref="DocxSession.SetHeaderFooterKindEnabled"/>.
    /// </summary>
    required public bool EvenAndOddHeaders { get; init; }

    /// <summary>Number of text columns. Defaults to 1 if no <c>w:cols</c> is set.</summary>
    required public int Columns { get; init; }

    /// <summary>URIs of the header parts referenced by this section, in declaration order.</summary>
    required public IReadOnlyList<string> HeaderPartUris { get; init; }

    /// <summary>URIs of the footer parts referenced by this section, in declaration order.</summary>
    required public IReadOnlyList<string> FooterPartUris { get; init; }

    /// <summary>
    /// The header stories that EFFECTIVELY apply to this section: its own
    /// <c>w:headerReference</c>s (in declaration order, each with its <c>w:type</c>) plus, for any
    /// kind it does not declare, the one it inherits from the nearest preceding section that does
    /// — flagged <see cref="HeaderFooterRef.Inherited"/>. This is what a renderer shows, so it is
    /// what a caller asking "which header applies here?" needs. <see cref="HeaderPartUris"/>
    /// remains this section's OWN references only.
    /// </summary>
    required public IReadOnlyList<HeaderFooterRef> HeaderRefs { get; init; }

    /// <summary>The footer stories that effectively apply to this section — see
    /// <see cref="HeaderRefs"/>.</summary>
    required public IReadOnlyList<HeaderFooterRef> FooterRefs { get; init; }

    /// <summary>
    /// The page number this section starts at (<c>w:pgNumType/@w:start</c>), or <c>null</c> when the
    /// attribute is absent — meaning the section continues the previous section's numbering. Read
    /// counterpart of <see cref="PageNumberingOp.Start"/>.
    /// </summary>
    public int? PageNumberStart { get; init; }

    /// <summary>
    /// This section's page-number format (<c>w:pgNumType/@w:fmt</c>), or <c>null</c> when the
    /// attribute is absent — meaning Word's default (<c>1, 2, 3</c>). Deliberately NOT defaulted to
    /// <see cref="NumberFormat.Decimal"/>: a UI needs to tell "inherits the default" from
    /// "explicitly decimal" to avoid writing an attribute the document never had.
    /// </summary>
    public NumberFormat? PageNumberFormat { get; init; }
}

/// <summary>
/// Snapshot of the high-signal "is this template fillable yet?" state for a
/// <see cref="DocxSession"/>. Returned by <see cref="DocxSession.GetEditSummary"/>.
/// Composes existing primitives — <see cref="DocxSession.FindPlaceholders"/>,
/// <see cref="DocxSession.Grep"/>, and the projection's <c>AnchorIndex</c> — into
/// a single struct so an agent can ask "what's left to fill in?" without
/// stitching three separate calls together.
/// </summary>
/// <remarks>
/// All counts are derived from the live document state at the moment the
/// summary is taken; mutate-then-read is the expected pattern. The placeholder
/// and underscore lists are disjoint by construction (the underscore regex
/// excludes runs already enclosed in <c>[…]</c>), so totaling them gives a
/// true count of remaining slots without double-counting.
/// </remarks>
public sealed record EditSummary
{
    /// <summary>Total number of anchors in the projection (paragraphs, headings,
    /// list items, tables, cells, footnotes, comments) — a rough proxy for
    /// document complexity / addressable surface.</summary>
    public int TotalAnchors { get; init; }

    /// <summary>Bracketed placeholders still present. Populated using
    /// <see cref="ProjectionScopes.All"/> — body + headers/footers/footnotes/endnotes/comments —
    /// so verification doesn't miss placeholders in non-body parts. Use
    /// <see cref="DocxSession.FindPlaceholders"/> directly for narrower scope.
    /// Empty when the template is fully filled.</summary>
    public IReadOnlyList<TemplatePlaceholder> RemainingPlaceholders { get; init; }
        = Array.Empty<TemplatePlaceholder>();

    /// <summary>Bare <c>___</c> runs of three or more underscores NOT enclosed in
    /// brackets — the second-class placeholder shape that <see cref="DocxSession.FindPlaceholders"/>
    /// deliberately skips. Surfaces here so callers see "fillable blanks Word
    /// authors sometimes leave outside brackets" without a manual <see cref="DocxSession.Grep"/>.</summary>
    public IReadOnlyList<TextMatch> BareUnderscoreRuns { get; init; }
        = Array.Empty<TextMatch>();

    /// <summary>Number of user-authored footnotes (excludes the two Word-reserved
    /// boilerplate notes: <c>w:type="separator"</c> and <c>w:type="continuationSeparator"</c>).</summary>
    public int FootnoteCount { get; init; }

    /// <summary>Number of inline <c>w:footnoteReference</c> markers in the main body —
    /// how many times any footnote is cited. May differ from <see cref="FootnoteCount"/>
    /// if a footnote is referenced multiple times or an orphan footnote exists.</summary>
    public int InlineFootnoteRefCount { get; init; }

    /// <summary>Number of comment anchors in the projection (excludes the comment
    /// range markers; counts each distinct comment thread once).</summary>
    public int CommentCount { get; init; }
}

/// <summary>How far below the target anchor to include in <see cref="DocxSession.ProjectAnchor"/>.</summary>
public enum ProjectionDepth
{
    /// <summary>Just the target block itself (its anchor + its own text). For headings,
    /// returns only the heading paragraph, not the section under it.</summary>
    SelfOnly = 0,

    /// <summary>Self + descendants. Most useful for <c>tbl</c> anchors (returns the whole
    /// table); for paragraphs it's the same as <see cref="SelfOnly"/>.</summary>
    Subtree = 1,

    /// <summary>Self + descendants + following siblings up to (but not including) the
    /// next sibling at the same or higher heading level. For non-heading anchors,
    /// equivalent to <see cref="Subtree"/>. This is the dominant "give me this section"
    /// case for headings and is the default.</summary>
    SubtreeAndFollowingSiblings = 2,
}

/// <summary>
/// Output format for <see cref="DocxSession.GetDiff(DiffFormat)"/>.
/// </summary>
public enum DiffFormat
{
    /// <summary>JSON array of <see cref="DiffEntry"/> records. The agentic-friendly
    /// shape — anchor-keyed, ordered by document position. Default.</summary>
    Json = 0,

    /// <summary>Standard unified diff (git-style) over the initial vs. current
    /// markdown projection. Line-based LCS; 3 lines of context per hunk; uses
    /// <c>--- initial</c> / <c>+++ current</c> as filename headers. Output is
    /// parseable by <c>patch(1)</c>. Empty string when nothing has changed.</summary>
    Unified = 1,

    /// <summary>Two-column human-review diff (<c>diff -y</c> style) over the
    /// initial vs. current markdown projection. Each row pairs an initial-side
    /// line with a current-side line; the centre column carries one of
    /// <c>' '</c> (unchanged), <c>'|'</c> (modified — both columns have content),
    /// <c>'&lt;'</c> (only initial — deleted), <c>'&gt;'</c> (only current —
    /// inserted). Left column is wrapped/padded to 72 chars.</summary>
    SideBySide = 2,
}

/// <summary>
/// A single anchor-keyed change in the diff between an initial and current projection.
/// </summary>
public sealed record DiffEntry
{
    /// <summary>Op kind: <c>"delete"</c> (anchor existed initially, gone now),
    /// <c>"insert"</c> (anchor exists now but not initially), or
    /// <c>"modify"</c> (anchor exists in both but with different content).</summary>
    required public string Op { get; init; }

    /// <summary>The anchor's id (current id for insert/modify; initial id for delete).</summary>
    required public string AnchorId { get; init; }

    /// <summary>Pre-change text content for delete/modify. <c>null</c> for insert.</summary>
    public string? Before { get; init; }

    /// <summary>Post-change text content for insert/modify. <c>null</c> for delete.</summary>
    public string? After { get; init; }
}

/// <summary>
/// The projection change one mutation produced (<see cref="EditResult.Patch"/>). Block-scoped: it
/// carries the markdown of the top-level blocks the edit changed, not the whole document (issue
/// #1022).
/// </summary>
/// <remarks>
/// <para>A top-level block is a child of the body, of a header or footer, a note definition, or a
/// comment, addressed by its anchor id. <see cref="Blocks"/> lists every block that changed or was
/// added since the previous patch, in document order, each with its current markdown and the block
/// it now follows; <see cref="RemovedAnchorIds"/> lists blocks that are gone. A client keeping one
/// entry per block applies a patch by deleting the removed ids, then, block by block, deleting the
/// id if present and re-inserting it after <see cref="MarkdownPatchBlock.AfterAnchorId"/>.</para>
/// <para>When a block-local patch cannot be exact the patch is the whole document:
/// <see cref="IsFullDocument"/> is true and <see cref="Markdown"/> is the full projection. That
/// happens for an edit that renumbered a list, changed a style, edited a header or footer, or added,
/// removed or renumbered a story part or a note; for the first patch after an op failed and rolled
/// back, a preview was committed, or the session's index had to be rebuilt; and always under a
/// non-<see cref="AnchorIdRendering.FullUnid"/> rendering.</para>
/// <para><see cref="DocxSession.Undo"/> and <see cref="DocxSession.Redo"/> return no patch: a
/// client mirroring the projection re-reads it (<see cref="DocxSession.Project"/>) after them, and
/// the next patch covers only what changes afterwards.</para>
/// </remarks>
/// <param name="ScopeAnchorId">The anchor the op addressed (its smallest enclosing block).</param>
/// <param name="Markdown">The changed blocks' markdown concatenated in document order, or the whole
/// projection when <see cref="IsFullDocument"/>.</param>
public sealed record MarkdownPatch(string ScopeAnchorId, string Markdown)
{
    /// <summary>True when <see cref="Markdown"/> is the whole projection rather than changed blocks;
    /// <see cref="Blocks"/> and <see cref="RemovedAnchorIds"/> are then empty.</summary>
    public bool IsFullDocument { get; init; }

    /// <summary>Blocks changed or added since the previous patch, in document order.</summary>
    public IReadOnlyList<MarkdownPatchBlock> Blocks { get; init; } = Array.Empty<MarkdownPatchBlock>();

    /// <summary>Anchor ids of blocks removed since the previous patch (or whose id changed, such as
    /// a paragraph that became a heading; the new id appears in <see cref="Blocks"/>). May name a
    /// block that was added and removed again in between, which a client simply ignores.</summary>
    public IReadOnlyList<string> RemovedAnchorIds { get; init; } = Array.Empty<string>();
}

/// <summary>One changed top-level block in a <see cref="MarkdownPatch"/>.</summary>
/// <param name="AnchorId">The block's anchor id (<c>kind:scope:unid</c>).</param>
/// <param name="AfterAnchorId">The anchor id of the block it now follows in its scope, or null when
/// it is the scope's first block.</param>
/// <param name="Markdown">The block's own markdown, without the separators the full projection puts
/// between blocks; empty for a block that renders nothing.</param>
public sealed record MarkdownPatchBlock(string AnchorId, string? AfterAnchorId, string Markdown);

/// <summary>One top-level render unit in a <see cref="RenderPlan"/> — a body block
/// (<c>p</c>/<c>h</c>/<c>li</c>), one whole table (<c>tbl</c>, its rows/cells/cell
/// paragraphs subsumed), or one footnote/endnote definition (<c>fn</c>/<c>en</c>).
/// <para><see cref="Sig"/> is a content signature carried ONLY by container units
/// (<c>tbl</c>/<c>fn</c>/<c>en</c>): a container's unid is structural (tag-name
/// signature) and survives edits INSIDE it — a row insert or a note text edit keeps
/// the container's unid — so a renderer diffing by unid alone would keep a stale
/// node. The signature hashes the descendant unids, which any inner content or
/// structure change re-derives, so a changed container diffs as an in-place
/// substitution. <c>null</c> for leaf blocks, whose own unid IS their content
/// signature.</para></summary>
public sealed record RenderUnit(string Id, string Kind, string? Sig = null, int Section = 0, int Group = 0);

/// <summary>
/// A block a move source may legally land against, and on which side. The two positions are
/// reported separately because a cross-block range or a section break BETWEEN the blocks can
/// make one side legal and the other not — see <see cref="DocxSession.ValidMoveTargets"/>.
/// </summary>
public sealed record MoveTarget(string AnchorId, bool Before, bool After);

/// <summary>
/// The ordered top-level render units per scope container — the authority for "what
/// blocks exist, in what order" that an incremental renderer diffs its DOM against.
/// The projection's flat <c>AnchorIndex</c> cannot express table containment (a cell
/// paragraph and a body paragraph are both kind <c>p</c>); this plan can.
/// </summary>
public sealed record RenderPlan(
    System.Collections.Generic.IReadOnlyList<RenderUnit> Body,
    System.Collections.Generic.IReadOnlyList<RenderUnit> Footnotes,
    System.Collections.Generic.IReadOnlyList<RenderUnit> Endnotes);

/// <summary>
/// One footnote/endnote in citation order. <see cref="Id"/> is the note's
/// <c>w:id</c> as written in the XML; <see cref="Ordinal"/> is its 1-based
/// citation position — which IS its displayed number (ids ascend in reference
/// order, the invariant every Word file holds). A client that renumbers rendered
/// note chrome (markers, hrefs, list values) after an insert walks its markers in
/// document order and applies the k-th entry to the k-th marker.
/// </summary>
public sealed record NoteListEntry(string Id, string DefAnchorId, int Ordinal);

/// <summary>
/// One native Word comment, in comments-part order — see <see cref="DocxSession.ListComments"/>.
/// <see cref="DefAnchorId"/> addresses the definition (kind <c>cmt</c>) for
/// <see cref="DocxSession.UpdateComment"/>/<see cref="DocxSession.RemoveComment"/>;
/// <see cref="Date"/> is the raw <c>w:date</c> attribute string (null when absent);
/// <see cref="Text"/> is the flattened body (paragraphs joined by a space, the
/// <c>w:annotationRef</c> mark excluded). <see cref="ParentAnchorId"/> resolves
/// <c>w15:paraIdParent</c> back to the parent definition anchor; <see cref="Resolved"/>
/// reflects <c>w15:done</c>. Both are null when this comment has no
/// <c>commentsExtended</c> entry. <see cref="Id"/> is the numeric <c>w:id</c> — the value the
/// HTML renderer stamps as <c>data-comment-id</c> on highlight spans and markers, which is how a
/// rendered range is matched back to this entry. Mutations still address comments by anchor.
/// </summary>
public sealed record CommentListEntry(
    string DefAnchorId, string Author, string? Initials, string? Date, string Text)
{
    /// <summary>The comment's numeric <c>w:comment/@w:id</c> (the same id the range markers and
    /// the rendered <c>data-comment-id</c> attribute carry).</summary>
    required public int Id { get; init; }

    /// <summary>The parent definition's stable <c>cmt</c> anchor for a reply; null for a
    /// top-level or legacy comment.</summary>
    public string? ParentAnchorId { get; init; }

    /// <summary>Word's <c>w15:done</c> state; null when no extension entry exists.</summary>
    public bool? Resolved { get; init; }
}

public enum RevisionFamily
{
    ContentInsert,
    ContentDelete,
    Move,
    ParagraphMark,
    RowInsert,
    RowDelete,
    CellInsert,
    CellDelete,
    CellMerge,
    ContentControlInsert,
    ContentControlDelete,
    NumberingPropertiesInsert,
    NumberingChange,
    PropertiesChange,
    Unsupported,
}

public enum RevisionResolutionStatus
{
    Supported,
    Unsupported,
    Malformed,
    Ambiguous,
}

public sealed record RevisionDiagnostic(string Code, string Message);

/// <summary>
/// One part-qualified, markup-native tracked revision. <see cref="AffectedAnchors"/>
/// contains every currently addressable structure the atomic revision can change;
/// <see cref="AnchorId"/> remains the convenient first/primary anchor.
/// Unsupported, malformed, and ambiguous markup is deliberately listed and fails
/// closed when resolution is requested.
/// </summary>
public sealed record RevisionListEntry
{
    required public string Id { get; init; }
    required public string Type { get; init; }
    required public RevisionFamily Family { get; init; }
    required public IReadOnlyList<string> ConstituentIds { get; init; }
    /// <summary>
    /// QName-qualified native carrier identities. Unlike <see cref="ConstituentIds"/>, these
    /// distinguish revision roles which legally use the same numeric <c>w:id</c> value.
    /// </summary>
    required public IReadOnlyList<string> ConstituentKeys { get; init; }
    required public string Author { get; init; }
    public string? Date { get; init; }
    public string? DateUtc { get; init; }
    required public string Text { get; init; }
    required public string PartUri { get; init; }
    required public string Scope { get; init; }
    public string? AnchorId { get; init; }
    required public IReadOnlyList<Anchor> AffectedAnchors { get; init; }
    required public RevisionResolutionStatus ResolutionStatus { get; init; }
    public RevisionDiagnostic? Diagnostic { get; init; }
}

/// <summary>The explicit repairs the revision registry can perform on markup it refuses to
/// resolve (issues #754–#758). Ordinary listing, accept and reject never repair.</summary>
public enum RevisionRepairKind
{
    /// <summary>Give every carrier of the entry a fresh document-unique <c>w:id</c>; markers that
    /// shared one old id keep sharing the new one. For missing, non-numeric, and duplicated ids.</summary>
    AssignIdentity,

    /// <summary>Move an orphan numbering marker into its paragraph's own <c>w:numPr</c>.</summary>
    ReattachNumberingChange,

    /// <summary>Move an orphan cell marker into the <c>w:tcPr</c> of its enclosing cell.</summary>
    ReattachCellMarker,

    /// <summary>Treat orphan deleted text as live text (<c>w:delText</c> → <c>w:t</c>).</summary>
    RestoreOrphanText,

    /// <summary>Wrap the run holding orphan deleted text in a <c>w:del</c> with caller-supplied
    /// author and date, making it an ordinary deletion.</summary>
    WrapOrphanTextAsDeletion,
}

/// <summary>One repair the registry offers for one listed entry, and whether it can perform it.</summary>
public sealed record RevisionRepairProposal
{
    required public string RevisionId { get; init; }
    required public RevisionRepairKind Kind { get; init; }
    required public string PartUri { get; init; }
    /// <summary>The defect being repaired, as the revision listing reports it.</summary>
    required public RevisionDiagnostic Diagnostic { get; init; }
    /// <summary>Every native carrier the repair touches, as QName@element-path keys.</summary>
    required public IReadOnlyList<string> Carriers { get; init; }
    required public bool Repairable { get; init; }
    /// <summary>What the repair does, or why the package's evidence does not permit it.</summary>
    required public string Reason { get; init; }
    /// <summary>The request must supply <c>Author</c> and <c>Date</c>.</summary>
    public bool RequiresAuthorship { get; init; }
}

/// <summary>One repair to perform: the listed entry, the offered kind, and the review metadata
/// a kind that fabricates none on its own must be given.</summary>
public sealed record RevisionRepairRequest
{
    required public string RevisionId { get; init; }
    required public RevisionRepairKind Kind { get; init; }
    public string? Author { get; init; }
    public string? Date { get; init; }
}

/// <summary>One carrier's identity before and after a repair. <see cref="NewId"/> is empty when the
/// repair moved the carrier without renumbering it.</summary>
public sealed record RevisionCarrierIdentity(string Carrier, string? OldId, string NewId);

public sealed record RevisionRepairOutcome
{
    required public string RevisionId { get; init; }
    required public RevisionRepairKind Kind { get; init; }
    required public string PartUri { get; init; }
    required public IReadOnlyList<RevisionCarrierIdentity> Identities { get; init; }
}

/// <summary>Result of <see cref="DocxSession.RepairRevisions"/>: atomic, one undo step. The public
/// ids of repaired entries change (they derive from carrier identity), so re-list afterwards.</summary>
public sealed record RevisionRepairResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public IReadOnlyList<RevisionRepairOutcome> Repairs { get; init; } = Array.Empty<RevisionRepairOutcome>();
    public IReadOnlyList<Anchor> Modified { get; init; } = Array.Empty<Anchor>();

    internal static RevisionRepairResult Fail(EditErrorCode code, string message) =>
        new() { Success = false, Error = new EditError(code, message) };
}

/// <summary>Summary returned by <see cref="DocxSession.CompactRuns"/>.</summary>
public sealed record CompactResult
{
    /// <summary>Number of <c>w:r</c> elements whose only content was a <c>w:rPr</c>
    /// (or which had no children at all) and were therefore removed. <c>0</c>
    /// means the document was already compact across the selected scopes.</summary>
    public int RunsRemoved { get; init; }
}

/// <summary>The current state of a precondition target, returned even when it no longer exists.</summary>
public sealed record PreconditionTarget
{
    public bool Exists { get; init; }
    public string? AnchorId { get; init; }
    public string? Kind { get; init; }
    public string? Scope { get; init; }
    public string? ContentHash { get; init; }
    public string? VisibleText { get; init; }
}

/// <summary>Structured expected/actual detail for <see cref="EditErrorCode.PreconditionFailed"/>.</summary>
public sealed record PreconditionFailure(
    string Condition,
    object? Expected,
    object? Actual,
    long CurrentVersion,
    PreconditionTarget? CurrentTarget);

/// <summary>
/// Optimistic guards evaluated immediately before a mutation. Anchor-specific fields use
/// <see cref="AnchorId"/> as their target; a stale kind prefix still resolves by Unid, just like
/// ordinary mutation addressing. <see cref="ExpectedTextRange"/> is measured against the exact
/// <see cref="AnchorInfo.VisibleText"/> value.
/// </summary>
public sealed record MutationPreconditions
{
    public long? ExpectedVersion { get; init; }
    public string? AnchorId { get; init; }
    public string? ExpectedContentHash { get; init; }
    public string? ExpectedText { get; init; }
    public TextRangePrecondition? ExpectedTextRange { get; init; }
    public string? ExpectedKind { get; init; }
    public string? ExpectedScope { get; init; }
    public int? ExpectedMatchCount { get; init; }
}

/// <summary>An exact substring assertion within an anchor's visible text.</summary>
public sealed record TextRangePrecondition(int Start, int Length, string Text);

/// <summary>Execution policy for a group of synchronous document mutations.</summary>
public enum MutationBatchMode
{
    /// <summary>All steps commit as one undo/version unit, or every step is rolled back.</summary>
    Atomic,

    /// <summary>Run every step independently and retain successful steps after failures.</summary>
    BestEffort,
}

/// <summary>
/// One core batch step. <see cref="Preflight"/> performs any read-only validation that can be
/// decided before mutation begins (all steps up front for atomic mode, or immediately before each
/// step for best-effort mode); <see cref="Mutation"/> returns zero or more edit envelopes. An
/// empty collection is a successful no-op, while multiple results let multi-match replacement
/// remain one step without losing its individual outcomes.
/// </summary>
public sealed class MutationBatchStep
{
    public MutationBatchStep(
        string tool,
        string action,
        Func<DocxSession, IReadOnlyList<EditResult>> mutation,
        Func<DocxSession, EditError?>? preflight = null)
    {
        Tool = tool ?? throw new ArgumentNullException(nameof(tool));
        Action = action ?? throw new ArgumentNullException(nameof(action));
        Mutation = mutation ?? throw new ArgumentNullException(nameof(mutation));
        Preflight = preflight;
    }

    public MutationBatchStep(
        string tool,
        string action,
        Func<DocxSession, EditResult> mutation,
        Func<DocxSession, EditError?>? preflight = null)
        : this(tool, action, s => new[] { mutation(s) }, preflight)
    {
    }

    public string Tool { get; }
    public string Action { get; }
    public Func<DocxSession, IReadOnlyList<EditResult>> Mutation { get; }
    public Func<DocxSession, EditError?>? Preflight { get; }

    /// <summary>
    /// The step's request arguments as a JSON object, for delivery evidence (issue #748). The
    /// delegate hides them, so a transport that knows them attaches them here; a step without
    /// them is recorded with empty arguments.
    /// </summary>
    public string? ArgumentsJson { get; init; }
}

/// <summary>Result of one batch step, including whether its effects were rolled back.</summary>
public sealed record MutationBatchStepResult(
    int Index,
    string Tool,
    string Action,
    IReadOnlyList<EditResult> Results,
    bool RolledBack)
{
    public bool Success => Results.All(r => r.Success);
}

/// <summary>The first failed step in a batch.</summary>
public sealed record MutationBatchFailure(
    int Index,
    string Tool,
    string Action,
    EditError Error,
    bool RolledBack);

/// <summary>Added, removed, and modified semantic objects predicted or produced by a batch.</summary>
public sealed record MutationBatchChangeSet<T>(
    IReadOnlyList<T> Added,
    IReadOnlyList<T> Removed,
    IReadOnlyList<T> Modified)
{
    public static MutationBatchChangeSet<T> Empty { get; } = new(
        Array.Empty<T>(), Array.Empty<T>(), Array.Empty<T>());
}

/// <summary>Optional HTML projection generated only from an isolated preview session.</summary>
public enum MutationPreviewHtmlMode
{
    None,
    Scoped,
    Full,
}

/// <summary>Optional outputs for <see cref="DocxSession.PreviewBatch"/>.</summary>
public sealed record MutationBatchPreviewOptions
{
    public MutationPreviewHtmlMode HtmlMode { get; init; }
    public string? HtmlAnchorId { get; init; }

    /// <summary>
    /// Keep a successful preview's exact result package so <see cref="DocxSession.CommitPreview"/>
    /// can later make it the live document with the previewed generated ids, timestamps and
    /// package hash (issue #760). The receipt then carries <see cref="MutationBatchResult.Retention"/>.
    /// Off by default: retention holds a whole package per preview.
    /// </summary>
    public bool Retain { get; init; }
}

/// <summary>
/// Identity of a preview retained for a guarded commit (issue #760). The commit is bound to the
/// live state the preview was predicted from: <see cref="BaseVersion"/> and
/// <see cref="BasePackageHash"/> must still describe the session, and the entry is gone after
/// <see cref="ExpiresAt"/>, after eviction, after the session closes, or once it is committed.
/// </summary>
public sealed record MutationPreviewRetention(
    string PreviewId,
    long BaseVersion,
    string BasePackageHash,
    DateTimeOffset ExpiresAt);

/// <summary>Structured result of an atomic or explicit best-effort mutation batch.</summary>
public sealed record MutationBatchResult
{
    public MutationBatchMode Mode { get; init; }
    public bool Preview { get; init; }
    public bool Success { get; init; }
    public bool RolledBack { get; init; }
    public long BaseVersion { get; init; }
    public long ResultVersion { get; init; }
    /// <summary>
    /// SHA-256 over ordered OPC entry names and uncompressed payload bytes. ZIP compression,
    /// timestamps, and entry framing are deliberately excluded; XML payload timestamps remain.
    /// A deterministic replay at <see cref="BaseVersion"/> should produce this hash. Generated
    /// anchors/OOXML ids and execution timestamps make other batches only semantically equivalent;
    /// callers must consult <see cref="Warnings"/> before using the hash as a replay assertion.
    ///
    /// <c>null</c> — never an empty string — when the hash could not be computed (the reason is
    /// in <see cref="Warnings"/>). An absent hash must not compare equal to another absent hash:
    /// a sentinel that does turns <c>preview.PackageHash == applied.PackageHash</c> into an
    /// assertion that passes precisely when it has nothing to assert.
    /// </summary>
    public string? PackageHash { get; init; }
    public IReadOnlyList<MutationBatchStepResult> Steps { get; init; } =
        Array.Empty<MutationBatchStepResult>();
    public MutationBatchFailure? Failure { get; init; }
    public MutationBatchChangeSet<RevisionListEntry> RevisionChanges { get; init; } =
        MutationBatchChangeSet<RevisionListEntry>.Empty;
    public MutationBatchChangeSet<CommentListEntry> CommentChanges { get; init; } =
        MutationBatchChangeSet<CommentListEntry>.Empty;
    public MutationBatchChangeSet<DocumentAnnotation> AnnotationChanges { get; init; } =
        MutationBatchChangeSet<DocumentAnnotation>.Empty;
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
    public string? Html { get; init; }

    /// <summary>
    /// Present on a preview retained for commit (<see cref="MutationBatchPreviewOptions.Retain"/>)
    /// and on the result of committing it; null for every other batch.
    /// </summary>
    public MutationPreviewRetention? Retention { get; init; }
}

/// <summary>
/// A synchronous, nested-safe document transaction. Dispose without <see cref="Commit"/> to
/// restore the complete package and all session/history state captured at begin. The scope owns
/// the session-wide mutation gate and must be completed on the thread that created it.
/// </summary>
public sealed class DocxSessionTransaction : IDisposable
{
    private DocxSession? _session;
    private readonly long _id;

    internal DocxSessionTransaction(DocxSession session, long id)
    {
        _session = session;
        _id = id;
    }

    public bool IsCompleted => _session is null || _session.IsDisposed;

    public void Commit()
    {
        var session = _session ?? throw new InvalidOperationException("transaction already completed");
        if (session.IsDisposed)
        {
            _session = null;
            throw new ObjectDisposedException(nameof(DocxSession));
        }
        session.CompleteTransaction(_id, commit: true);
        // Keep the scope recoverable if owner-thread/LIFO validation (or completion itself)
        // throws. CompleteTransaction only removes the state after a valid completion.
        _session = null;
    }

    public void Rollback()
    {
        var session = _session ?? throw new InvalidOperationException("transaction already completed");
        if (session.IsDisposed)
        {
            _session = null;
            throw new ObjectDisposedException(nameof(DocxSession));
        }
        session.CompleteTransaction(_id, commit: false);
        _session = null;
    }

    public void Dispose()
    {
        if (_session is null) return;
        var session = _session;
        // Disposing the owning session explicitly abandons and invalidates its active scopes.
        // A later using-scope unwind must be inert rather than reopening the disposed package.
        if (session.IsDisposed)
        {
            _session = null;
            return;
        }
        session.CompleteTransaction(_id, commit: false);
        _session = null;
    }
}

public sealed record EditError(EditErrorCode Code, string Message, string? AnchorId = null)
{
    public PreconditionFailure? Precondition { get; init; }
}

public enum EditErrorCode
{
    AnchorNotFound,
    AnchorWrongKind,
    AnchorsNotAdjacent,
    SessionDisposed,

    MalformedMarkdown,
    UnsupportedMarkdownSyntax,
    TableInsertNotSupported,
    FootnoteRefNotSupported,
    CommentMarkerNotSupported,
    ImageInsertNotSupported,
    AnchorTokenInPayload,

    OffsetOutOfRange,
    InvalidPosition,

    /// <summary>A text needle that matches nothing in the target anchor's visible text.
    /// <see cref="DocxSession.ReplaceTextRange"/> reports it instead of a silent
    /// zero-replacement so a caller can distinguish "replaced nothing" from "replaced
    /// something" — re-inspect the anchor and retarget. Pass
    /// <see cref="ReplaceOptions.ExpectedMatchCount"/> = 0 to assert absence instead.</summary>
    TextNotFound,

    UnknownStyle,
    InvalidListLevel,

    /// <summary>A list start value OOXML cannot express: <c>w:startOverride/@w:val</c> is a
    /// non-negative decimal, so <see cref="DocxSession.SetListStartOverride"/> rejects a
    /// negative value.</summary>
    InvalidListStartValue,

    /// <summary>A page-numbering value that OOXML cannot express: a start page below zero, or
    /// <see cref="NumberFormat.Bullet"/> as a page-number format (neither <c>w:pgNumType/@w:fmt</c>
    /// nor the field <c>\*</c> switch has a bullet notion).</summary>
    InvalidPageNumbering,

    /// <summary>A page-setup value the section cannot express: a non-positive page width/height,
    /// a negative margin or header/footer distance, opposing margins that leave no room for
    /// content (<c>left + right &gt;= width</c>, <c>top + bottom &gt;= height</c>), or
    /// <see cref="DocxSession.SetHeaderFooterKindEnabled"/> asked to switch off
    /// <see cref="HeaderFooterKind.Default"/>, which has no on/off flag.</summary>
    InvalidPageSetup,

    /// <summary>A reference-field option is out of range or malformed (issue #607).</summary>
    InvalidReferenceField,

    /// <summary>A <see cref="ParagraphFormatOp"/> that OOXML cannot express: both
    /// <c>FirstLineIndent</c> and <c>HangingIndent</c> in one op (<c>w:ind</c> holds one or the
    /// other), a negative indent/spacing value (the attributes are unsigned), or a
    /// <c>LineSpacingRule</c> without the <c>LineSpacing</c> it qualifies.</summary>
    InvalidParagraphFormat,

    /// <summary>A table-styling value the op cannot express: a column-width list whose length
    /// doesn't match the table's column count (or a non-positive width), a shading fill that is
    /// neither a hex RRGGBB triplet nor "auto", or a negative border size.</summary>
    InvalidTableStyling,

    /// <summary>A cell merge/unmerge the grid cannot express: a rectangle running past the table's
    /// last row/column, a rectangle whose rows do not tile the same whole grid columns (it would
    /// partially overlap an existing <c>w:gridSpan</c>), a rectangle that clips a vertical merge
    /// entering from above or continuing below, a merge covering fewer than two cells, absorbed
    /// content under <see cref="TableMergeContent.Reject"/>, or an unmerge of a cell that carries
    /// no merge markup.</summary>
    InvalidTableMerge,

    /// <summary>The supplied anchor is not the canonical <c>tc</c> cell anchor and cannot be
    /// translated by the compatibility shim. During the compatibility window only a legacy
    /// paragraph/heading/list-item anchor physically inside the intended cell is translated;
    /// use table metadata or coordinate resolution to obtain the cell's <c>tc</c> anchor.</summary>
    TableAnchorMigrationRequired,

    MalformedXml,
    DisallowedNamespace,
    IncompatibleElementType,
    ValidationFailed,

    NothingToUndo,
    NothingToRedo,

    DuplicateAnnotationId,
    AnnotationNotFound,
    EmptyAnnotationSpan,

    HyperlinkNotFound,
    BookmarkNotFound,
    DuplicateBookmarkName,
    InvalidBookmarkName,
    InvalidHyperlinkTarget,
    MissingBookmarkTarget,
    BookmarkInUse,
    ManagedBookmark,
    EmptyHyperlinkSpan,
    UnsupportedInlineBoundary,

    ImageNotFound,
    InvalidImageData,
    UnsupportedImageFormat,
    ImageTooLarge,
    InvalidImageDimensions,
    UnsupportedImageMarkup,
    LinkedImageReadOnly,
    InvalidImageLayout,

    ContentControlNotFound,
    ContentControlMalformed,
    ContentControlUnsupported,
    ContentControlLocked,
    ContentControlBound,
    ContentControlWrongType,
    InvalidContentControlValue,
    ContentControlPlacementUnsupported,
    ContentControlNestedFillUnsupported,
    RepeatingSectionConstraint,

    /// <summary>A zero-length span passed to <see cref="DocxSession.AddComment"/>, or a
    /// whole-block comment requested on a paragraph with no text — a comment range must
    /// cover at least one character.</summary>
    EmptyCommentSpan,

    /// <summary>The revision id passed to <see cref="DocxSession.AcceptRevision"/>/
    /// <see cref="DocxSession.RejectRevision"/> matches no revision in the current
    /// markup — never listed, already resolved, or removed by resolving an enclosing
    /// revision. Re-<see cref="DocxSession.ListRevisions"/> for the current set.</summary>
    RevisionNotFound,

    /// <summary>An optimistic mutation guard did not match the current session or target state.</summary>
    PreconditionFailed,

    /// <summary>A mutation batch step names an unsupported operation or a read-only action.</summary>
    InvalidBatchStep,

    /// <summary>A transaction identity was supplied where it cannot safely identify an applying batch.</summary>
    InvalidTransaction,

    /// <summary>A transaction id was already reserved for a different canonical request.</summary>
    TransactionConflict,

    /// <summary>The exact response for a known transaction was evicted from bounded retention.</summary>
    TransactionResultEvicted,

    /// <summary>A known transaction never recorded a terminal response, so its outcome is unknown.</summary>
    TransactionIncomplete,

    /// <summary>No retained preview has this id: it was never retained, expired, was evicted,
    /// or was already committed.</summary>
    PreviewNotFound,

    /// <summary>The retained preview was predicted from a live state the session no longer has
    /// (version, package content, tracked-changes mode or revision author changed).</summary>
    PreviewStale,

    /// <summary>The revision family is visible but has no safe selective resolver.</summary>
    RevisionUnsupported,

    /// <summary>The native marker topology is incomplete or internally inconsistent.</summary>
    RevisionMalformed,

    /// <summary>The native identity/topology maps to more than one possible operation.</summary>
    RevisionAmbiguous,

    /// <summary>A requested revision repair is not one the registry offers for that entry, is
    /// not repairable from the package's evidence, or lacks the authorship it requires.</summary>
    RevisionRepairRejected,

    /// <summary>The requested mutation has no reversible native tracked-change encoding.</summary>
    TrackedOperationUnsupported,

    /// <summary>A structural edit was refused because its table still has unresolved structure revisions.</summary>
    UnresolvedStructuralRevision,

    InternalError,

    /// <summary>A failed op's rollback also failed (<see cref="DocxSession.LastRollbackError"/>), so
    /// the document may be half-mutated. Every later mutation is refused with this code until the
    /// session is closed and reopened from known-good bytes (issue #963).</summary>
    SessionCorrupted,
}

public sealed class EditResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public IReadOnlyList<Anchor> Created { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Removed { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Modified { get; init; } = Array.Empty<Anchor>();
    public MarkdownPatch? Patch { get; init; }

    /// <summary>Structural identity mapping populated by table shape mutations.</summary>
    public TableAnchorMapping? TableAnchors { get; init; }

    /// <summary>
    /// Populated by AddAnnotation/RemoveAnnotation/UpdateAnnotation/MoveAnnotation
    /// with the affected annotation id. Null for every other op.
    /// </summary>
    public string? AnnotationId { get; init; }

    /// <summary>The affected native hyperlink/bookmark identity, when applicable.</summary>
    public string? HyperlinkId { get; init; }
    public string? BookmarkName { get; init; }

    /// <summary>The affected native image occurrence identity, when applicable.</summary>
    public string? ImageId { get; init; }

    internal static EditResult Fail(EditErrorCode code, string message, string? anchorId = null) =>
        new() { Success = false, Error = new EditError(code, message, anchorId) };

    internal static EditResult Fail(EditError error) => new() { Success = false, Error = error };
}

/// <summary>
/// Partial-update payload for <see cref="DocxSession.UpdateAnnotation"/>.
/// Null fields leave the existing value unchanged. <see cref="MetadataPatch"/>
/// is a per-key merge: a non-null value sets the key, an explicit null removes
/// it, a missing key leaves it unchanged.
/// </summary>
public sealed record AnnotationUpdate
{
    public string? LabelId { get; init; }
    public string? Label { get; init; }
    public string? Color { get; init; }
    public string? Author { get; init; }
    public IReadOnlyDictionary<string, string?>? MetadataPatch { get; init; }
}

public sealed class DocxSessionSettings
{
    /// <summary>
    /// Maximum number of undo steps retained. Lowered from 50 to 20 — a depth alone never
    /// bounded memory, because each step is a deep clone of every snapshot-scoped part and so
    /// costs whatever the DOCUMENT costs. Fifty steps over a long filing is fifty whole-document
    /// DOMs held live. Raise it freely for small documents; <see cref="UndoMemoryBudgetBytes"/>
    /// is the bound that actually protects the heap.
    /// </summary>
    public int UndoDepth { get; init; } = 20;

    /// <summary>
    /// Approximate ceiling, in bytes, on the memory retained by undo/redo snapshots. When
    /// exceeded the ring discards its OLDEST entries (redo first, then undo) until it is back
    /// under budget, so the depth cap and this budget are both upper bounds and whichever binds
    /// first wins.
    ///
    /// <para>Default 128 MiB. Set to 0 (or any non-positive value) for the historical behavior of
    /// depth-only bounding — appropriate for a server process editing modest documents, and a
    /// latent OOM in a browser WASM heap, which is why it is not the default.</para>
    ///
    /// <para>The measure is an ESTIMATE of retained heap, not serialized size: a live LINQ-to-XML
    /// tree costs several times the XML it came from. One undo step is always retained even if a
    /// single snapshot exceeds the whole budget, and <see cref="DocxSession.UndoHistoryTrimmedForMemory"/>
    /// reports whether the budget has ever discarded history.</para>
    /// </summary>
    public long UndoMemoryBudgetBytes { get; init; } = 128L * 1024 * 1024;
    public bool ValidateRawOps { get; init; } = false;
    public TrackedChangeMode TrackedChanges { get; init; } = TrackedChangeMode.Accept;
    public string? RevisionAuthor { get; init; }
    public WmlToMarkdownConverterSettings ProjectionSettings { get; init; } = new();

    /// <summary>
    /// When <c>false</c> (default) <see cref="DocxSession.Save"/> strips
    /// <c>PtOpenXml:Unid</c> attributes from every part — the attribute is internal
    /// to the projector and not in the OOXML schema, so persisting it bloats saved
    /// DOCX files (a 100-page document grows by ~700 KB of attribute noise). Set to
    /// <c>true</c> when anchor ids must survive a save/reopen round trip — the
    /// scenario flagged by Open Question #1 in <c>docs/architecture/markdown_projection.md</c>.
    /// </summary>
    public bool PersistAnchorIds { get; init; } = false;

    /// <summary>
    /// When <c>true</c>, <c>ReplaceText</c>/<c>ReplaceTextRange</c>/<c>ReplaceMatch</c>/
    /// <c>ReplaceTextAtSpan</c>/<c>ReplaceInner</c> payloads (and replacements passed to
    /// <c>InsertParagraph</c> / <c>ReplaceCellContent</c>)
    /// have ASCII <c>"</c> and <c>'</c> converted to typographic curly quotes
    /// (U+201C/U+201D and U+2018/U+2019) based on context — open quote at the start
    /// of a string, after whitespace, or after an open-bracket; close quote elsewhere.
    /// Avoids the cosmetic regression where a replacement lands as <c>"foo"</c> next
    /// to surrounding <c>"foo"</c> already-curly text. Default <c>false</c> (pass payloads
    /// through unchanged) — see issue #140.
    /// </summary>
    public bool SmartQuotes { get; init; } = false;

    /// <summary>
    /// When <c>false</c>, mutation ops return <c>Patch = null</c> and skip the per-op
    /// scope re-projection that builds it. For clients that re-render from HTML (the
    /// browser editor) the patch is dead weight — on a 350-block document it is a large
    /// share of every op's latency. Default <c>true</c> (wire-compatible).
    /// </summary>
    public bool EmitMarkdownPatch { get; init; } = true;

    /// <summary>
    /// When <c>true</c> (default), the session projects the document at construction
    /// time and stashes the result so <see cref="DocxSession.GetDiff"/> can compare
    /// initial vs. current. It also retains an exact copy of the opening package for
    /// <see cref="DocxSession.GetSemanticChanges"/>. Costs ~200ms plus one package copy
    /// at construction for a 100-page doc; turn off to skip the upfront time/memory when
    /// you don't plan to call either comparison API.
    /// </summary>
    public bool CaptureInitialProjection { get; init; } = true;

    /// <summary>
    /// Record the evidence a delivery change receipt needs as edits execute (issue #748): the
    /// exact package before and after every version step, the request each transport described,
    /// transaction identities, and undo/redo lineage. Requires <see cref="CaptureInitialProjection"/>
    /// (the opening package is the receipt's source document). Off by default: every version
    /// step then serializes a clean package copy, and retention holds them until delivery.
    /// </summary>
    public bool CaptureDeliveryEvidence { get; init; } = false;

    /// <summary>
    /// Verification-only mode: resolution may remove only artifacts attributable to the selected
    /// revision. Ordinary editing also scopes relationship cleanup to the resolved changes,
    /// but retains its historical handling of other empty markup containers.
    /// </summary>
    internal bool ProofSafeRevisionResolution { get; init; }

    /// <summary>Numbering instances present in the proof path's expected baseline.</summary>
    internal IReadOnlyList<int> ProtectedRevisionNumberingIds { get; init; } = Array.Empty<int>();

    /// <summary>Abstract numbering definitions present in the proof path's expected endpoint.</summary>
    internal IReadOnlyList<int> ProtectedRevisionAbstractNumberingIds { get; init; }
        = Array.Empty<int>();

    /// <summary>
    /// Owner-part/relationship-id pairs present in the proof path's expected endpoint. A
    /// generated revision may reuse an otherwise orphaned endpoint relationship; resolving the
    /// revision must not erase that pre-existing package state.
    /// </summary>
    internal IReadOnlyCollection<string> ProtectedRevisionRelationshipKeys { get; init; }
        = Array.Empty<string>();

    /// <summary>
    /// Empty optional property shells present in the proof path's expected endpoint. Resolution
    /// may remove a selected revision's now-empty direct shell only when that shell kind is not
    /// represented in the endpoint. This is deliberately conservative: exact package comparison
    /// remains the final guard when multiple shells of the same kind exist.
    /// </summary>
    internal IReadOnlyCollection<string> ProtectedRevisionEmptyContainerKeys { get; init; }
        = Array.Empty<string>();
}
