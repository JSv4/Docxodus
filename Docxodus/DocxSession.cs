#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

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
/// <see cref="TotalPages"/> → a <c>NUMPAGES</c> field (the total page count).
/// </summary>
public enum PageNumberField { CurrentPage, TotalPages }

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

/// <summary>Which table edges a <see cref="DocxSession.SetTableBorders"/> call targets:
/// <see cref="Outside"/> = top/left/bottom/right, <see cref="Inside"/> = the inner grid lines
/// (<c>w:insideH</c>/<c>w:insideV</c>), <see cref="All"/> = both.</summary>
public enum TableBorderScope { All, Outside, Inside }

/// <summary>Shading granularity for <see cref="DocxSession.SetCellShading"/>: the one cell the
/// anchor sits in, or every cell of its row (header-row banding).</summary>
public enum TableShadingScope { Cell, Row }

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
    /// <summary>The Unid of the <c>w:sectPr</c> element this info describes. Stable across mutations.</summary>
    required public string SectionUnid { get; init; }

    required public int PageWidthTwips { get; init; }
    required public int PageHeightTwips { get; init; }
    required public bool Landscape { get; init; }
    required public int MarginTopTwips { get; init; }
    required public int MarginBottomTwips { get; init; }
    required public int MarginLeftTwips { get; init; }
    required public int MarginRightTwips { get; init; }

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

public sealed record MarkdownPatch(string ScopeAnchorId, string Markdown);

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
public sealed record RenderUnit(string Id, string Kind, string? Sig = null);

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
/// <c>commentsExtended</c> entry. The numeric <c>w:id</c> is deliberately not surfaced —
/// comments are addressed by anchor everywhere in this API.
/// </summary>
public sealed record CommentListEntry(
    string DefAnchorId, string Author, string? Initials, string? Date, string Text)
{
    /// <summary>The parent definition's stable <c>cmt</c> anchor for a reply; null for a
    /// top-level or legacy comment.</summary>
    public string? ParentAnchorId { get; init; }

    /// <summary>Word's <c>w15:done</c> state; null when no extension entry exists.</summary>
    public bool? Resolved { get; init; }
}

/// <summary>
/// One tracked revision, read directly off the live document's markup in document
/// order — see <see cref="DocxSession.ListRevisions"/>. <see cref="Id"/> is stable
/// while the underlying markup exists (derived from the markup's own <c>w:id</c>
/// attributes, so resolving OTHER revisions never renames it) and is what
/// <see cref="DocxSession.AcceptRevision"/>/<see cref="DocxSession.RejectRevision"/>
/// address. <see cref="Type"/> is <c>"insert"</c>, <c>"delete"</c>, <c>"move"</c>
/// (a linked move pair — both sides resolve together), or <c>"format"</c>.
/// <see cref="Author"/>/<see cref="Date"/> are the true <c>w:author</c>/<c>w:date</c>
/// from the markup (date null when absent). <see cref="Text"/> is the revision's
/// visible text (the deleted text for deletions, <c>¶</c> for a revised paragraph
/// mark, the affected text for format changes). <see cref="AnchorId"/> is the
/// containing block's anchor (null when the block isn't projection-addressable).
/// </summary>
public sealed record RevisionListEntry(
    string Id, string Type, string Author, string? Date, string Text, string? AnchorId);

/// <summary>Summary returned by <see cref="DocxSession.CompactRuns"/>.</summary>
public sealed record CompactResult
{
    /// <summary>Number of <c>w:r</c> elements whose only content was a <c>w:rPr</c>
    /// (or which had no children at all) and were therefore removed. <c>0</c>
    /// means the document was already compact across the selected scopes.</summary>
    public int RunsRemoved { get; init; }
}

public sealed record EditError(EditErrorCode Code, string Message, string? AnchorId = null);

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

    /// <summary>A <see cref="ParagraphFormatOp"/> that OOXML cannot express: both
    /// <c>FirstLineIndent</c> and <c>HangingIndent</c> in one op (<c>w:ind</c> holds one or the
    /// other), a negative indent/spacing value (the attributes are unsigned), or a
    /// <c>LineSpacingRule</c> without the <c>LineSpacing</c> it qualifies.</summary>
    InvalidParagraphFormat,

    /// <summary>A table-styling value the op cannot express: a column-width list whose length
    /// doesn't match the table's column count (or a non-positive width), a shading fill that is
    /// neither a hex RRGGBB triplet nor "auto", or a negative border size.</summary>
    InvalidTableStyling,

    MalformedXml,
    DisallowedNamespace,
    IncompatibleElementType,
    ValidationFailed,

    NothingToUndo,
    NothingToRedo,

    DuplicateAnnotationId,
    AnnotationNotFound,
    EmptyAnnotationSpan,

    /// <summary>A zero-length span passed to <see cref="DocxSession.AddComment"/>, or a
    /// whole-block comment requested on a paragraph with no text — a comment range must
    /// cover at least one character.</summary>
    EmptyCommentSpan,

    /// <summary>The revision id passed to <see cref="DocxSession.AcceptRevision"/>/
    /// <see cref="DocxSession.RejectRevision"/> matches no revision in the current
    /// markup — never listed, already resolved, or removed by resolving an enclosing
    /// revision. Re-<see cref="DocxSession.ListRevisions"/> for the current set.</summary>
    RevisionNotFound,

    InternalError,
}

public sealed class EditResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public IReadOnlyList<Anchor> Created { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Removed { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Modified { get; init; } = Array.Empty<Anchor>();
    public MarkdownPatch? Patch { get; init; }

    /// <summary>
    /// Populated by AddAnnotation/RemoveAnnotation/UpdateAnnotation/MoveAnnotation
    /// with the affected annotation id. Null for every other op.
    /// </summary>
    public string? AnnotationId { get; init; }

    internal static EditResult Fail(EditErrorCode code, string message, string? anchorId = null) =>
        new() { Success = false, Error = new EditError(code, message, anchorId) };
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
    public int UndoDepth { get; init; } = 50;
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
    /// When <c>true</c>, <c>ReplaceText</c>/<c>ReplaceTextRange</c>/<c>ReplaceMatch</c>
    /// payloads (and replacements passed to <c>InsertParagraph</c> / <c>ReplaceCellContent</c>)
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
    /// initial vs. current. Costs ~200ms at construction for a 100-page doc; turn
    /// off to skip the upfront cost when you don't plan to call <c>GetDiff</c>.
    /// </summary>
    public bool CaptureInitialProjection { get; init; } = true;
}

// ─── Session ───────────────────────────────────────────────────────────────

public sealed class DocxSession : IDisposable
{
    private readonly DocxSessionSettings _settings;
    private readonly Internal.UndoRing<DocumentSnapshot> _history;
    private MemoryStream? _stream;
    private WordprocessingDocument? _doc;
    private MarkdownProjection? _cachedProjection;
    private MarkdownProjection? _initialProjection;
    private bool _disposed;
    private int _revisionCounter = 1000;
    private long _lastFormatRevisionTicks;
    private RawDocxOps? _raw;

    // Mutable session configuration (issue #304): seeded from _settings at construction,
    // switchable mid-session via SetTrackedChanges/SetRevisionAuthor. Session config, not
    // document state — never captured in undo snapshots.
    private TrackedChangeMode _trackedChanges;
    private string? _revisionAuthor;

    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        _settings = settings ?? new DocxSessionSettings();
        _trackedChanges = _settings.TrackedChanges;
        _revisionAuthor = _settings.RevisionAuthor;
        _history = new Internal.UndoRing<DocumentSnapshot>(_settings.UndoDepth);
        _stream = new MemoryStream();
        _stream.Write(docxBytes, 0, docxBytes.Length);
        _stream.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);

        if (_settings.CaptureInitialProjection)
            _initialProjection = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
    }

    public Exception? LastInternalError { get; private set; }

    /// <summary>How subsequent mutations are recorded — switchable mid-session (issue #304).</summary>
    public TrackedChangeMode TrackedChanges => _trackedChanges;

    /// <summary>Author stamped on tracked-change markup; null means the "docxodus" default.</summary>
    public string? RevisionAuthor => _revisionAuthor;

    /// <summary>
    /// Switch how subsequent mutations are recorded. Session configuration, not a document
    /// mutation: takes no undo snapshot (Undo/Redo never change the mode) and never touches
    /// already-applied markup — switching to <see cref="TrackedChangeMode.Accept"/> does not
    /// accept existing revisions, and switching to <see cref="TrackedChangeMode.RenderInline"/>
    /// does not retroactively wrap prior direct edits.
    /// </summary>
    public void SetTrackedChanges(TrackedChangeMode mode)
    {
        if (_disposed) return;
        _trackedChanges = mode;
    }

    /// <summary>
    /// Change the author stamped on subsequent tracked-change markup (null restores the
    /// "docxodus" default). Session configuration — same non-undoable semantics as
    /// <see cref="SetTrackedChanges"/>. Enables multi-author edits in one session.
    /// </summary>
    public void SetRevisionAuthor(string? author)
    {
        if (_disposed) return;
        _revisionAuthor = author;
    }

    public MarkdownProjection Project()
    {
        ThrowIfDisposed();
        return _cachedProjection ??=
            WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
    }

    /// <summary>
    /// Project a sub-region of the document anchored at <paramref name="anchorId"/>.
    /// Returns a <see cref="MarkdownProjection"/> whose <c>Markdown</c> contains only
    /// the blocks in scope (per <paramref name="depth"/>) and whose <c>AnchorIndex</c>
    /// is filtered to those blocks plus their descendants.
    /// </summary>
    /// <param name="anchorId">The anchor to project. Must exist in the current
    /// <see cref="Project"/>'s AnchorIndex.</param>
    /// <param name="depth">How far below the target to include. Default
    /// <see cref="ProjectionDepth.SubtreeAndFollowingSiblings"/> — for headings, returns
    /// the full section bounded by the next same-or-higher heading.</param>
    /// <returns>A <see cref="MarkdownProjection"/> scoped to the requested region.</returns>
    /// <exception cref="InvalidOperationException">If <paramref name="anchorId"/> isn't in the AnchorIndex.</exception>
    public MarkdownProjection ProjectAnchor(
        string anchorId,
        ProjectionDepth depth = ProjectionDepth.SubtreeAndFollowingSiblings)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);

        var fullProjection = Project();
        var target = FindAnchor(anchorId)
            ?? throw new InvalidOperationException($"anchor not found: {anchorId}");

        var startElement = target.Resolve(_doc!)
            ?? throw new InvalidOperationException($"anchor element resolved null: {anchorId}");

        // Compute the set of Unids in scope.
        var inRange = new HashSet<string>(StringComparer.Ordinal);
        CollectUnids(startElement, inRange);

        if (depth == ProjectionDepth.SubtreeAndFollowingSiblings && target.Anchor.Kind == "h")
        {
            // For headings, also include forward siblings up to next same-or-higher heading.
            int targetLevel = WmlToMarkdownConverter.HeadingLevel(startElement);
            foreach (var sibling in startElement.ElementsAfterSelf())
            {
                if (sibling.Name == W.p
                    && WmlToMarkdownConverter.IsHeading(sibling)
                    && WmlToMarkdownConverter.HeadingLevel(sibling) <= targetLevel)
                {
                    break;  // hit the section boundary
                }
                CollectUnids(sibling, inRange);
            }
        }
        else if (depth == ProjectionDepth.Subtree)
        {
            // CollectUnids already added self + descendants; nothing more to do.
        }

        // SelfOnly: descendants shouldn't be in scope — keep just the starting element's Unid.
        if (depth == ProjectionDepth.SelfOnly)
        {
            inRange.Clear();
            var selfUnid = (string?)startElement.Attribute(PtOpenXml.Unid);
            if (selfUnid is not null) inRange.Add(selfUnid);
        }

        // Filter the full markdown to blocks whose anchor token is in-range.
        // Blocks are separated by blank lines; each in-range block starts with {#kind:scope:unid}.
        var sb = new System.Text.StringBuilder();
        foreach (var block in fullProjection.Markdown.Split("\n\n"))
        {
            var match = System.Text.RegularExpressions.Regex.Match(block, @"\{#[^:]+:[^:]+:([^\s}]+)\}");
            if (!match.Success) continue;  // skip scope markers / dividers / etc.
            // The rendered id might be the abbreviated or sequential form — translate back
            // to the full Unid via the dual-keyed AnchorIndex.
            if (TryResolveToUnid(match, fullProjection, out var fullUnid)
                && inRange.Contains(fullUnid))
            {
                sb.Append(block).Append("\n\n");
            }
        }

        // Filter the AnchorIndex too — keep only entries whose Unid is in scope.
        var filteredIndex = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
        foreach (var (key, value) in fullProjection.AnchorIndex)
        {
            if (inRange.Contains(value.Unid))
                filteredIndex[key] = value;
        }

        return new MarkdownProjection
        {
            Markdown = sb.ToString().TrimEnd('\n'),
            AnchorIndex = filteredIndex,
        };
    }

    private static void CollectUnids(XElement el, HashSet<string> sink)
    {
        var unid = (string?)el.Attribute(PtOpenXml.Unid);
        if (unid is not null) sink.Add(unid);
        foreach (var d in el.Descendants())
        {
            var dUnid = (string?)d.Attribute(PtOpenXml.Unid);
            if (dUnid is not null) sink.Add(dUnid);
        }
    }

    /// <summary>
    /// Resolve a rendered anchor id (full Unid, abbreviation, or sequential) back to
    /// the underlying full Unid by looking it up in the projection's AnchorIndex
    /// (which is dual-keyed when AnchorIdRendering is Abbreviated/Sequential).
    /// </summary>
    private static bool TryResolveToUnid(
        System.Text.RegularExpressions.Match match,
        MarkdownProjection projection,
        out string fullUnid)
    {
        // The full key is the content between {# and } — works for FullUnid and as an
        // alias key for Abbreviated/Sequential modes (BuildAnchorIndex dual-keys the index).
        var fullKey = match.Value.Substring(2, match.Value.Length - 3);
        if (projection.AnchorIndex.TryGetValue(fullKey, out var target))
        {
            fullUnid = target.Unid;
            return true;
        }
        fullUnid = match.Groups[1].Value;
        return false;
    }

    /// <summary>
    /// Looks up an anchor id with a fallback to Unid-only resolution. The dictionary
    /// is keyed by full <c>kind:scope:unid</c> id, so when a mutation flips the kind
    /// prefix (e.g., <c>p:body:abcd</c> → <c>h:body:abcd</c> after promoting to a
    /// heading), a cached old id would otherwise miss. This helper trails through
    /// to a Unid scan, so agents that hold cached ids keep working — matching the
    /// promise in <c>docs/architecture/docx_mutation_api.md</c>.
    /// </summary>
    /// <summary>
    /// The anchor index for LOOKUP (mutations, EditResult anchors). Reuses the full
    /// projection's index when one is cached; otherwise builds and caches the cheap
    /// index-only variant (no markdown emission, no per-entry TextPreview/AutoNumberPrefix)
    /// — see <see cref="WmlToMarkdownConverter.BuildAnchorIndexOnly"/>. Entries from this
    /// path therefore carry empty previews; consumers that need enrichment must call
    /// <see cref="Project"/> explicitly.
    /// </summary>
    internal IReadOnlyDictionary<string, AnchorTarget> AnchorIndex()
    {
        ThrowIfDisposed();
        if (_cachedProjection is not null) return _cachedProjection.AnchorIndex;
        return _cachedAnchorIndex ??=
            WmlToMarkdownConverter.BuildAnchorIndexOnly(_doc!, _settings.ProjectionSettings);
    }

    private IReadOnlyDictionary<string, AnchorTarget>? _cachedAnchorIndex;

    /// <summary>
    /// The ordered top-level render units per scope container — see <see cref="RenderPlan"/>.
    /// Body = the main body's direct children in document order (each <c>w:p</c> under its
    /// projected kind, each <c>w:tbl</c> as ONE <c>tbl</c> unit); Footnotes/Endnotes = the
    /// non-boilerplate note definitions in part order. Elements the projection does not
    /// address (e.g. <c>w:sectPr</c>) are skipped. Unlike the projection itself, empty
    /// paragraphs are ALWAYS listed — the plan mirrors the rendered DOM, which contains
    /// every block regardless of <see cref="EmptyParagraphMode"/>.
    /// </summary>
    public RenderPlan ListBlocks()
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids are assigned on every projected part

        var body = new List<RenderUnit>();
        var bodyEl = _doc!.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (bodyEl is not null)
        {
            foreach (var el in bodyEl.Elements())
            {
                string? kind =
                    el.Name == W.tbl ? "tbl" :
                    el.Name == W.p ? WmlToMarkdownConverter.KindFor(el) : null;
                var unid = (string?)el.Attribute(PtOpenXml.Unid);
                if (kind is null || unid is null) continue;
                body.Add(new RenderUnit($"{kind}:body:{unid}", kind, UnidHelper.ContentHash(el)));
            }
        }

        List<RenderUnit> Notes(XElement? root, XName noteName, bool endnotes, string kindScope)
        {
            var list = new List<RenderUnit>();
            if (root is null) return list;
            // MIRRORS THE RENDERER exactly (WmlToHtmlConverter's notes sections), which
            // is the only contract that lets a DOM diff work:
            //  - with ≥1 citation, the section renders the CITED notes in citation order
            //    (the tracker path) — an uncited note (Word's continuationNotice) does
            //    NOT render;
            //  - with zero citations, it renders every non-separator note in part order
            //    (so an uncited notice DOES render there).
            var cited = ListNotes(endnotes);
            if (cited.Count > 0)
            {
                foreach (var n in cited)
                    list.Add(new RenderUnit(n.DefAnchorId, kindScope,
                        ResolveNoteDef(root, noteName, n.Id) is { } def ? UnidHelper.ContentHash(def) : null));
                return list;
            }
            foreach (var n in root.Elements(noteName))
            {
                if ((string?)n.Attribute(W.type) is "separator" or "continuationSeparator") continue;
                var unid = (string?)n.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                list.Add(new RenderUnit($"{kindScope}:{kindScope}:{unid}", kindScope, UnidHelper.ContentHash(n)));
            }
            return list;
        }

        static XElement? ResolveNoteDef(XElement root, XName noteName, string id) =>
            root.Elements(noteName).FirstOrDefault(n => (string?)n.Attribute(W.id) == id);

        var main = _doc!.MainDocumentPart;
        return new RenderPlan(
            body,
            Notes(main?.FootnotesPart?.GetXDocument().Root, W.footnote, endnotes: false, "fn"),
            Notes(main?.EndnotesPart?.GetXDocument().Root, W.endnote, endnotes: true, "en"));
    }

    /// <summary>
    /// The document's footnotes (or endnotes) in citation order — see
    /// <see cref="NoteListEntry"/>. References are collected from the main body in
    /// document order (Word disallows note references anywhere else this API can
    /// author them); a reference whose definition is missing is skipped.
    /// </summary>
    public IReadOnlyList<NoteListEntry> ListNotes(bool endnotes = false)
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids on the note parts

        var result = new List<NoteListEntry>();
        var main = _doc!.MainDocumentPart;
        var bodyRoot = main?.GetXDocument().Root;
        var notesRoot = endnotes
            ? main?.EndnotesPart?.GetXDocument().Root
            : main?.FootnotesPart?.GetXDocument().Root;
        if (bodyRoot is null || notesRoot is null) return result;

        var refName = endnotes ? W.endnoteReference : W.footnoteReference;
        var defName = endnotes ? W.endnote : W.footnote;
        var kindScope = endnotes ? "en" : "fn";

        var defsById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var def in notesRoot.Elements(defName))
        {
            var defId = (string?)def.Attribute(W.id);
            if (defId is not null) defsById[defId] = def;
        }

        foreach (var r in bodyRoot.Descendants(refName))
        {
            var id = (string?)r.Attribute(W.id);
            if (id is null || !defsById.TryGetValue(id, out var def)) continue;
            var unid = (string?)def.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            result.Add(new NoteListEntry(id, $"{kindScope}:{kindScope}:{unid}", result.Count + 1));
        }
        return result;
    }

    /// <summary>
    /// Remove a native Word comment, addressed by its definition anchor (kind <c>cmt</c>):
    /// the <c>w:comment</c> definition, its body-side marker triple
    /// (<c>w:commentRangeStart</c>/<c>w:commentRangeEnd</c>/<c>w:commentReference</c>, wrapper
    /// run included) everywhere in the package, and any <c>commentsExtended</c>/
    /// <c>commentsIds</c> threading entries keyed by its paragraphs' <c>w14:paraId</c> — a
    /// surviving reply whose parent was removed becomes top-level. Delegates to the same
    /// teardown <see cref="DeleteBlock"/> performs for a <c>cmt</c> anchor; this wrapper adds
    /// only the comment-specific kind guard. The comments part itself is kept even when the
    /// last comment is removed (part deletion happens only via <see cref="Undo"/> of the
    /// create).
    /// </summary>
    public EditResult RemoveComment(string commentAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(commentAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {commentAnchorId}", commentAnchorId);
        if (target.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"RemoveComment requires a comment definition anchor (kind cmt); got kind={target.Anchor.Kind}",
                commentAnchorId);

        return DeleteBlock(commentAnchorId);
    }

    /// <summary>
    /// The document's native Word comments in comments-part order — see
    /// <see cref="CommentListEntry"/>. Read-only; returns an empty list when the document
    /// has no comments part.
    /// </summary>
    public IReadOnlyList<CommentListEntry> ListComments()
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids on the comments part

        var result = new List<CommentListEntry>();
        var main = _doc!.MainDocumentPart;
        var root = main?.WordprocessingCommentsPart?.GetXDocument().Root;
        if (root is null) return result;

        var comments = root.Elements(W.comment).ToList();
        var anchorByParaId = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var c in comments)
        {
            var unid = (string?)c.Attribute(PtOpenXml.Unid);
            var paraId = (string?)c.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
            if (unid is not null && paraId is not null)
                anchorByParaId[paraId] = $"cmt:cmt:{unid}";
        }

        var commentExByParaId = main?.WordprocessingCommentsExPart?.GetXDocument().Root?
            .Elements(Internal.CommentOps.W15 + "commentEx")
            .Where(e => (string?)e.Attribute(Internal.CommentOps.W15 + "paraId") is not null)
            .GroupBy(e => (string)e.Attribute(Internal.CommentOps.W15 + "paraId")!, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal)
            ?? new Dictionary<string, XElement>(StringComparer.Ordinal);

        foreach (var c in comments)
        {
            var unid = (string?)c.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;

            string? parentAnchorId = null;
            bool? resolved = null;
            var paraId = (string?)c.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
            if (paraId is not null && commentExByParaId.TryGetValue(paraId, out var commentEx))
            {
                resolved = Internal.CommentOps.ParseDone(
                    (string?)commentEx.Attribute(Internal.CommentOps.W15 + "done"));
                var parentParaId = (string?)commentEx.Attribute(Internal.CommentOps.W15 + "paraIdParent");
                if (parentParaId is not null)
                    anchorByParaId.TryGetValue(parentParaId, out parentAnchorId);
            }

            result.Add(new CommentListEntry(
                $"cmt:cmt:{unid}",
                (string?)c.Attribute(W.author) ?? "unknown",
                (string?)c.Attribute(W.initials),
                (string?)c.Attribute(W.date),
                Internal.CommentOps.FlattenBodyText(c))
            {
                ParentAnchorId = parentAnchorId,
                Resolved = resolved,
            });
        }
        return result;
    }

    // ─── Tracked revisions: markup-native listing + selective resolution (issue #318) ───

    /// <summary>
    /// Enumerate the document's tracked revisions directly off the live markup, in
    /// document order across every story RevisionProcessor walks (body, headers,
    /// footers, footnotes, endnotes). Contiguous markup of the same kind and author
    /// groups into one entry per user-visible change (an inserted paragraph is ONE
    /// revision: its runs plus its mark); a named move pair is one <c>"move"</c> entry
    /// covering both sides. Ids derive from the markup's <c>w:id</c> attributes, so
    /// they are stable across calls and across resolution of other revisions —
    /// unlike the re-diff listing, authors/dates are the markup's own. Not
    /// enumerated in v1 (still resolved by whole-document accept/reject):
    /// <c>cellIns</c>/<c>cellDel</c>/<c>cellMerge</c>, content-control ins/del
    /// ranges, and <c>numPr</c> numbering-ins markers.
    /// </summary>
    public IReadOnlyList<RevisionListEntry> ListRevisions()
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids so entries can carry block anchors

        var parts = RevisionStoryParts();
        var groups = Internal.RevisionOps.Enumerate(parts.Select(p => p.Root).ToList());
        var result = new List<RevisionListEntry>(groups.Count);
        foreach (var g in groups)
        {
            var partUri = parts[g.PartIndex].Part.Uri.ToString();
            string? anchorId = null;
            if (g.Units.Count > 0)
            {
                var first = g.Units[0];
                for (var a = first.Paragraph ?? first.MarkedRow ?? first.Element; a is not null; a = a.Parent)
                {
                    var unid = (string?)a.Attribute(PtOpenXml.Unid);
                    if (unid is null) continue;
                    if (AnchorForUnid(unid, partUri) is { } anch) anchorId = anch.Id;
                    break;
                }
            }
            result.Add(new RevisionListEntry(
                g.Id, g.Type, g.Author, g.Date, Internal.RevisionOps.GroupText(g), anchorId));
        }
        return result;
    }

    /// <summary>Accept ONE revision by the id <see cref="ListRevisions"/> reported —
    /// insertions keep their content (markup unwrapped), deletions are carried out,
    /// a move materializes at its destination, a format change keeps the new
    /// properties. An undoable session mutation; every other revision's markup (and
    /// id) is left untouched.</summary>
    public EditResult AcceptRevision(string revisionId) => ResolveRevision(revisionId, accept: true);

    /// <summary>Reject ONE revision by id — the inverse of <see cref="AcceptRevision"/>:
    /// insertions are removed, deleted content is restored (<c>w:delText</c> back to
    /// <c>w:t</c>), a move stays at its source, a format change restores the stored
    /// old properties.</summary>
    public EditResult RejectRevision(string revisionId) => ResolveRevision(revisionId, accept: false);

    private EditResult ResolveRevision(string revisionId, bool accept)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (string.IsNullOrEmpty(revisionId))
            return EditResult.Fail(EditErrorCode.RevisionNotFound, "revision id is empty");

        _ = AnchorIndex();
        var parts = RevisionStoryParts();
        var groups = Internal.RevisionOps.Enumerate(parts.Select(p => p.Root).ToList());
        var group = groups.FirstOrDefault(x => x.Id == revisionId);
        if (group is null)
            return EditResult.Fail(EditErrorCode.RevisionNotFound, $"revision not found: {revisionId}");

        var partUri = parts[group.PartIndex].Part.Uri.ToString();

        // Capture the block anchors the resolution touches BEFORE applying — elements
        // detach during Apply and can no longer be resolved to a part afterwards.
        var modified = new List<Anchor>();
        var seenModified = new HashSet<string>(StringComparer.Ordinal);
        foreach (var u in group.Units)
        {
            for (var a = u.Paragraph ?? u.MarkedRow ?? u.Element; a is not null; a = a.Parent)
            {
                var unid = (string?)a.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                if (AnchorForUnid(unid, partUri) is { } anch && seenModified.Add(anch.Id))
                    modified.Add(anch);
                break;
            }
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var removedElements = Internal.RevisionOps.Apply(group, accept);

            var removed = new List<Anchor>();
            var seenRemoved = new HashSet<string>(StringComparer.Ordinal);
            foreach (var el in removedElements)
            {
                foreach (var d in el.DescendantsAndSelf())
                {
                    var unid = (string?)d.Attribute(PtOpenXml.Unid);
                    if (unid is null) continue;
                    if (AnchorForUnid(unid, partUri) is { } anch && seenRemoved.Add(anch.Id))
                        removed.Add(anch);
                }
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = modified.Where(m => !seenRemoved.Contains(m.Id)).ToList(),
                Removed = removed,
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>The story parts revision markup lives in, in the fixed order the
    /// revision enumeration indexes them (main, headers, footers, footnotes, endnotes
    /// — the same set RevisionProcessor's whole-document accept/reject walks).</summary>
    private List<(OpenXmlPart Part, XElement Root)> RevisionStoryParts()
    {
        var list = new List<(OpenXmlPart, XElement)>();
        foreach (var part in EnumerateProjectedPartsForScopes(
            ProjectionScopes.Body | ProjectionScopes.Headers | ProjectionScopes.Footers
            | ProjectionScopes.Footnotes | ProjectionScopes.Endnotes))
        {
            var root = part.GetXDocument().Root;
            if (root is not null) list.Add((part, root));
        }
        return list;
    }

    internal AnchorTarget? FindAnchor(string? anchorId)
    {
        if (anchorId is null) return null;
        var index = AnchorIndex();
        if (index.TryGetValue(anchorId, out var direct)) return direct;
        int lastColon = anchorId.LastIndexOf(':');
        if (lastColon <= 0 || lastColon == anchorId.Length - 1) return null;
        var unid = anchorId.Substring(lastColon + 1);
        foreach (var v in index.Values)
        {
            if (v.Unid == unid) return v;
        }
        return null;
    }

    /// <summary>
    /// Reverse-resolve an <see cref="Anchor"/> from the current projection by Unid, preferring
    /// the entry that lives in <paramref name="preferPartUri"/>.
    /// </summary>
    /// <remarks>
    /// Unids are CONTENT-ADDRESSED, so identical content in DIFFERENT package parts yields the
    /// SAME unid — a document with empty default/first/even header stories has one unid shared
    /// across several header parts (Word writes exactly that). A bare-unid reverse lookup then
    /// returns whichever part the projection happened to index first, so an <see cref="EditResult"/>
    /// would report an anchor pointing at the WRONG story, and a caller that addressed the
    /// returned anchor next (as an editor does) would silently write into the wrong part.
    /// Scoping by the part the edit actually touched keeps the round trip unambiguous; the
    /// unid-only fallback preserves behavior when the part isn't known.
    /// </remarks>
    private Anchor? AnchorForUnid(string? unid, string? preferPartUri)
    {
        if (unid is null) return null;
        AnchorTarget? fallback = null;
        foreach (var t in AnchorIndex().Values)
        {
            if (t.Unid != unid) continue;
            if (preferPartUri is not null && t.PartUri == preferPartUri) return t.Anchor;
            fallback ??= t;
        }
        return fallback?.Anchor;
    }

    /// <summary>URI of the package part that owns <paramref name="element"/>, or <c>null</c>
    /// when it belongs to no projected part (e.g. a detached element).</summary>
    private string? PartUriOf(XElement element)
    {
        var root = element.AncestorsAndSelf().Last();
        foreach (var part in EnumerateProjectedParts())
        {
            if (ReferenceEquals(part.GetXDocument().Root, root)) return part.Uri.ToString();
        }
        return null;
    }

    /// <summary>Reverse-resolve the anchor of a live element, scoped to its owning part.</summary>
    private Anchor? AnchorForElement(XElement element) =>
        AnchorForUnid((string?)element.Attribute(PtOpenXml.Unid), PartUriOf(element));

    public bool Exists(string anchorId)
    {
        ThrowIfDisposed();
        return FindAnchor(anchorId) is not null;
    }

    public AnchorInfo? GetAnchorInfo(string anchorId)
    {
        ThrowIfDisposed();
        _ = Project(); // AnchorInfo's product IS the enrichment — never serve the index-only (empty-preview) entries.
        var target = FindAnchor(anchorId);
        if (target is null) return null;
        return new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, target.TextPreview)
        {
            AutoNumberPrefix = target.AutoNumberPrefix,
        };
    }

    /// <summary>
    /// Bulk variant of <see cref="GetAnchorInfo"/>. Resolves every requested anchor
    /// from the projection's cached <c>AnchorIndex</c> in a single pass. Unknown
    /// anchor ids map to <c>null</c> in the returned dictionary so callers can
    /// distinguish "anchor doesn't exist" from "anchor exists with empty preview."
    /// </summary>
    public IReadOnlyDictionary<string, AnchorInfo?> GetAnchorInfos(IEnumerable<string> anchorIds)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorIds);
        _ = Project(); // See GetAnchorInfo — enrichment required, index-only entries won't do.

        var result = new Dictionary<string, AnchorInfo?>(StringComparer.Ordinal);
        foreach (var id in anchorIds)
        {
            if (id is null) continue;
            if (result.ContainsKey(id)) continue;
            var target = FindAnchor(id);
            result[id] = target is null
                ? null
                : new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, target.TextPreview)
                {
                    AutoNumberPrefix = target.AutoNumberPrefix,
                };
        }
        return result;
    }

    /// <summary>
    /// Resolves block-level metadata (style id + name, outline level, list
    /// membership, formatting probe) for <paramref name="anchorId"/>. Returns
    /// <c>null</c> when the anchor doesn't exist. See <see cref="BlockMetadata"/>
    /// for the field reference.
    /// </summary>
    public BlockMetadata? GetBlockMetadata(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetBlockMetadata(_doc!, target);
    }

    /// <summary>
    /// Bulk variant of <see cref="GetBlockMetadata"/>. Unknown anchor ids map
    /// to <c>null</c>; duplicate ids are deduped; iteration order matches
    /// input order for keys that appear first.
    /// </summary>
    public IReadOnlyDictionary<string, BlockMetadata?> GetBlockMetadatas(IEnumerable<string> anchorIds)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorIds);

        var result = new Dictionary<string, BlockMetadata?>(StringComparer.Ordinal);
        foreach (var id in anchorIds)
        {
            if (id is null) continue;
            if (result.ContainsKey(id)) continue;
            var target = FindAnchor(id);
            result[id] = target is null ? null : Internal.BlockMetadataOps.GetBlockMetadata(_doc!, target);
        }
        return result;
    }

    /// <summary>
    /// Resolves the <see cref="ListMembership"/> for a list-item paragraph;
    /// returns <c>null</c> when the anchor has no <c>w:numPr</c> (inline or
    /// inherited from style) or doesn't exist.
    /// </summary>
    public ListMembership? GetListMembership(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetListMembership(_doc!, target);
    }

    /// <summary>
    /// Resolves the <see cref="SectionInfo"/> for the <c>w:sectPr</c> that
    /// governs <paramref name="anchorId"/>. Returns <c>null</c> when the
    /// anchor lives outside the body part (footnotes, endnotes, headers,
    /// footers, comments) or doesn't exist.
    /// </summary>
    public SectionInfo? GetSectionInfo(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetSectionInfo(_doc!, target);
    }

    /// <summary>
    /// Searches the flat text of every paragraph/heading/list-item in <paramref name="scope"/>
    /// for matches of <paramref name="pattern"/> and returns them in document order, each
    /// with the run fragments it spans. The fragment list lets callers rewrite a match in
    /// place while preserving each fragment's formatting — see #143 for design context.
    /// </summary>
    /// <param name="pattern">Regular-expression pattern (use <c>Regex.Escape</c> for literal text).</param>
    /// <param name="options">Standard <see cref="System.Text.RegularExpressions.RegexOptions"/> flags.</param>
    /// <param name="scope">Which package parts to search. Defaults to <see cref="ProjectionScopes.Body"/>.</param>
    /// <param name="contextChars">Number of characters of surrounding text to include in
    /// <see cref="TextMatch.ContextBefore"/> and <see cref="TextMatch.ContextAfter"/>.</param>
    public IReadOnlyList<TextMatch> Grep(
        string pattern,
        System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 80,
        WhitespaceMode whitespace = WhitespaceMode.Preserve,
        ContextBoundary boundary = ContextBoundary.Char)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(pattern)) return Array.Empty<TextMatch>();

        var regex = new System.Text.RegularExpressions.Regex(pattern, options);
        var results = new List<TextMatch>();

        // Walk the projection's AnchorIndex so document order is the same order
        // an agent sees in the projection. Only block-level kinds that hold runs
        // qualify (paragraphs/headings/list-items/table cells); other kinds either
        // don't contain text directly (tbl, tr, sec) or live in non-body scopes
        // we filter explicitly below.
        var index = Project().AnchorIndex;
        foreach (var target in index.Values)
        {
            if (!ScopeMatches(target.Anchor.Scope, scope)) continue;
            if (target.Anchor.Kind is not ("p" or "h" or "li" or "tc")) continue;

            var element = target.Resolve(_doc!);
            if (element is null) continue;

            // Table cells contain paragraphs; recurse so a Grep over the body
            // also hits cell text. Other kinds operate on the element directly.
            if (target.Anchor.Kind == "tc")
            {
                // Cell paragraphs are reachable via their own AnchorIndex entries,
                // so skip the cell wrapper to avoid double-counting matches.
                continue;
            }

            var map = Internal.RunTextMap.Build(element);
            if (map.FlatText.Length == 0) continue;

            // Look up the owner part once per anchor so the hyperlink resolver
            // doesn't have to walk back up to the root annotation per run.
            var ownerPart = ResolvePart(target.PartUri);

            // For Normalize mode: match against a whitespace-normalized COPY of the
            // flat text while keeping the segment offset map pointing at the original
            // positions. Match indices apply unchanged because the substitutions are
            // 1:1 (NBSP → space, narrow-NBSP → space, thin-space → space) — same
            // character count, just different code points.
            var matchText = whitespace == WhitespaceMode.Normalize
                ? NormalizeWhitespace(map.FlatText)
                : map.FlatText;

            foreach (System.Text.RegularExpressions.Match m in regex.Matches(matchText))
            {
                if (!m.Success || m.Length == 0) continue;

                var pieces = Internal.RunTextMap.ResolveRange(map, m.Index, m.Length);
                if (pieces.Count == 0) continue;

                var fragments = new List<RunFragment>(pieces.Count);
                foreach (var (seg, offsetInRun, len) in pieces)
                {
                    var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
                    var runText = RunText(seg.Run);
                    fragments.Add(new RunFragment
                    {
                        Unid = runUnid,
                        Text = runText.Substring(offsetInRun, len),
                        SpanInElement = new CharSpan(offsetInRun, len),
                        Formatting = ExtractFormatting(seg.Run, ownerPart),
                    });
                }

                var (ctxBefore, ctxAfter) = WalkContext(map.FlatText, m.Index, m.Length, contextChars, boundary);

                var groups = new string[m.Groups.Count];
                for (int i = 0; i < m.Groups.Count; i++) groups[i] = m.Groups[i].Value;

                results.Add(new TextMatch
                {
                    Text = m.Value,
                    EnclosingAnchor = target,
                    Span = new CharSpan(m.Index, m.Length),
                    Fragments = fragments,
                    ContextBefore = ctxBefore,
                    ContextAfter = ctxAfter,
                    Groups = groups,
                });
            }
        }

        return results;
    }

    /// <summary>
    /// Searches the flat text of every block-level element in <paramref name="scope"/>, like
    /// <see cref="Grep"/>, but lets a single match span <em>adjacent</em> block-level siblings
    /// (paragraphs/headings/list items) sharing the same direct parent. Returns matches in
    /// document order, each with a per-block <see cref="BlockSlice"/> breakdown. See issue #146.
    ///
    /// Block boundaries are represented in the concatenated text by a single <c>\n</c>, so
    /// <c>^</c>/<c>$</c> with <see cref="System.Text.RegularExpressions.RegexOptions.Multiline"/>
    /// anchor at boundaries; <c>.</c> won't cross unless
    /// <see cref="System.Text.RegularExpressions.RegexOptions.Singleline"/> is set.
    ///
    /// Matches never cross:
    /// <list type="bullet">
    ///   <item><description>OOXML package parts (e.g. body → footnote, header → body).</description></item>
    ///   <item><description>Container boundaries (e.g. body paragraph → table-cell paragraph).</description></item>
    ///   <item><description>Non-paragraph siblings (a <c>w:tbl</c> or section property between two paragraphs breaks the run).</description></item>
    /// </list>
    ///
    /// Superset of <see cref="Grep"/>: single-block matches are still returned (with one
    /// <see cref="BlockSlice"/>). Callers that want only cross-block hits can filter
    /// <c>Slices.Count &gt; 1</c>.
    /// </summary>
    public IReadOnlyList<CrossBlockMatch> GrepCrossBlock(
        string pattern,
        System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 80,
        WhitespaceMode whitespace = WhitespaceMode.Preserve,
        ContextBoundary boundary = ContextBoundary.Char)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(pattern)) return Array.Empty<CrossBlockMatch>();

        var regex = new System.Text.RegularExpressions.Regex(pattern, options);
        var results = new List<CrossBlockMatch>();

        // Build groups of consecutive block-level siblings under the same parent.
        // Document order comes from AnchorIndex iteration; the parent check ensures
        // we don't bridge a body paragraph to a table-cell paragraph or a header to a
        // body paragraph. Any non-eligible anchor (kind != p/h/li, or out of scope,
        // or unresolved) breaks the run.
        var index = Project().AnchorIndex;
        var groups = new List<List<(AnchorTarget Target, XElement Element)>>();
        List<(AnchorTarget, XElement)>? current = null;
        XElement? currentParent = null;

        foreach (var target in index.Values)
        {
            if (!ScopeMatches(target.Anchor.Scope, scope)) { current = null; continue; }
            if (target.Anchor.Kind is not ("p" or "h" or "li")) { current = null; continue; }

            var element = target.Resolve(_doc!);
            if (element is null) { current = null; continue; }

            if (current is not null && ReferenceEquals(element.Parent, currentParent))
            {
                current.Add((target, element));
            }
            else
            {
                current = new List<(AnchorTarget, XElement)> { (target, element) };
                currentParent = element.Parent;
                groups.Add(current);
            }
        }

        foreach (var group in groups)
        {
            // Build per-block maps + a parallel boundary array (start offset of each
            // block in the concatenated text, length of the block's flat text). A
            // single '\n' between blocks acts as the sentinel.
            var maps = new List<Internal.RunTextMap.Map>(group.Count);
            var starts = new int[group.Count];
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < group.Count; i++)
            {
                if (i > 0) sb.Append('\n');
                starts[i] = sb.Length;
                var map = Internal.RunTextMap.Build(group[i].Element);
                maps.Add(map);
                sb.Append(map.FlatText);
            }
            var concat = sb.ToString();
            if (concat.Length == 0) continue;

            var matchText = whitespace == WhitespaceMode.Normalize
                ? NormalizeWhitespace(concat)
                : concat;

            // Cache owner-part lookup per group; every block in a group lives in the
            // same package part (siblings share a parent), so one lookup suffices.
            var ownerPart = ResolvePart(group[0].Target.PartUri);

            foreach (System.Text.RegularExpressions.Match m in regex.Matches(matchText))
            {
                if (!m.Success || m.Length == 0) continue;

                var slices = new List<BlockSlice>();
                var anchors = new List<AnchorTarget>();
                for (int i = 0; i < group.Count; i++)
                {
                    var blockStart = starts[i];
                    var blockEnd = blockStart + maps[i].FlatText.Length;
                    if (blockEnd <= m.Index) continue;
                    if (blockStart >= m.Index + m.Length) break;

                    var overlapStart = Math.Max(m.Index, blockStart) - blockStart;
                    var overlapLen = Math.Min(m.Index + m.Length, blockEnd) - blockStart - overlapStart;

                    var pieces = overlapLen > 0
                        ? Internal.RunTextMap.ResolveRange(maps[i], overlapStart, overlapLen)
                        : new List<(Internal.RunTextMap.RunSegment, int, int)>();

                    var fragments = new List<RunFragment>(pieces.Count);
                    foreach (var (seg, offsetInRun, len) in pieces)
                    {
                        var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
                        var runText = RunText(seg.Run);
                        fragments.Add(new RunFragment
                        {
                            Unid = runUnid,
                            Text = runText.Substring(offsetInRun, len),
                            SpanInElement = new CharSpan(offsetInRun, len),
                            Formatting = ExtractFormatting(seg.Run, ownerPart),
                        });
                    }

                    slices.Add(new BlockSlice
                    {
                        Anchor = group[i].Target,
                        SpanInBlock = new CharSpan(overlapStart, overlapLen),
                        Fragments = fragments,
                    });
                    anchors.Add(group[i].Target);
                }

                if (slices.Count == 0) continue;

                var (ctxBefore, ctxAfter) = WalkContext(concat, m.Index, m.Length, contextChars, boundary);

                var groups2 = new string[m.Groups.Count];
                for (int i = 0; i < m.Groups.Count; i++) groups2[i] = m.Groups[i].Value;

                results.Add(new CrossBlockMatch
                {
                    Text = m.Value,
                    EnclosingAnchors = anchors,
                    Slices = slices,
                    ContextBefore = ctxBefore,
                    ContextAfter = ctxAfter,
                    Groups = groups2,
                });
            }
        }

        return results;
    }

    /// <summary>
    /// Finds the first anchor whose flat text contains <paramref name="needle"/>, or null.
    /// Thin wrapper over <see cref="Grep"/> — every consumer was reimplementing the same
    /// scan with its own quirks (case sensitivity, NBSP, scope filter). See issue #137.
    /// </summary>
    public AnchorTarget? FindByText(string needle, FindOptions? options = null) =>
        FindAllByText(needle, options).FirstOrDefault();

    /// <summary>
    /// All anchors whose flat text contains <paramref name="needle"/>, in document order.
    /// Duplicates removed (one entry per enclosing anchor regardless of how many times
    /// the needle appears inside it).
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindAllByText(string needle, FindOptions? options = null)
    {
        if (string.IsNullOrEmpty(needle)) return Array.Empty<AnchorTarget>();
        var opts = options ?? new FindOptions();
        var regexOpts = opts.IgnoreCase
            ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
            : System.Text.RegularExpressions.RegexOptions.None;
        return FindMatchesFiltered(System.Text.RegularExpressions.Regex.Escape(needle), regexOpts, opts);
    }

    /// <summary>
    /// All anchors with at least one match for <paramref name="pattern"/>, in document order.
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindByRegex(
        string pattern,
        System.Text.RegularExpressions.RegexOptions regexOptions = System.Text.RegularExpressions.RegexOptions.None,
        FindOptions? options = null) =>
        FindMatchesFiltered(pattern, regexOptions, options ?? new FindOptions());

    /// <summary>
    /// All anchors of a given kind (and optionally scope), in document order. Direct read
    /// over the projection's <c>AnchorIndex</c>; no text scan, so no <see cref="FindOptions"/>.
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindByKind(string kind, string? scope = null)
    {
        ThrowIfDisposed();
        var result = new List<AnchorTarget>();
        foreach (var target in Project().AnchorIndex.Values)
        {
            if (target.Anchor.Kind != kind) continue;
            if (scope is not null && target.Anchor.Scope != scope) continue;
            result.Add(target);
        }
        return result;
    }

    private IReadOnlyList<AnchorTarget> FindMatchesFiltered(
        string pattern,
        System.Text.RegularExpressions.RegexOptions regexOptions,
        FindOptions options)
    {
        ThrowIfDisposed();
        // Prefer Scopes (typed, composable) for the underlying Grep walker. The
        // string ScopeFilter still applies as a finer post-filter below for
        // callers targeting a single named part like "hdr1".
        var matches = Grep(
            pattern,
            regexOptions,
            options.Scopes,
            contextChars: 0,
            whitespace: options.IgnoreWhitespace ? WhitespaceMode.Normalize : WhitespaceMode.Preserve);

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<AnchorTarget>();
        foreach (var m in matches)
        {
            var anchor = m.EnclosingAnchor;
            if (options.KindFilter is not null && anchor.Anchor.Kind != options.KindFilter) continue;
            if (options.ScopeFilter is not null && anchor.Anchor.Scope != options.ScopeFilter) continue;
            if (!seen.Add(anchor.Anchor.Id)) continue;
            result.Add(anchor);
        }
        return result;
    }

    /// <summary>
    /// Enumerate every anchor whose scope belongs to <paramref name="scopes"/>, in
    /// projection order. Convenience over walking <c>Project().AnchorIndex</c> and
    /// filtering by scope name — common for callers that want to operate on every
    /// header paragraph, every footnote, etc.
    /// </summary>
    /// <example>
    /// <code>
    /// // Every paragraph in any header or footer:
    /// foreach (var t in session.AnchorsByScope(ProjectionScopes.Headers | ProjectionScopes.Footers))
    ///     Console.WriteLine($"{t.Anchor.Scope}: {t.TextPreview}");
    /// </code>
    /// </example>
    public IReadOnlyList<AnchorTarget> AnchorsByScope(ProjectionScopes scopes)
    {
        ThrowIfDisposed();
        var result = new List<AnchorTarget>();
        foreach (var t in Project().AnchorIndex.Values)
            if (scopes.IncludesScope(t.Anchor.Scope))
                result.Add(t);
        return result;
    }

    // ─── Annotation-based anchor discovery (#132) ────────────────────────

    /// <summary>
    /// Resolves an annotation's range to the block-level markdown anchors covering it,
    /// in document order. The bridge between the read-side annotation API
    /// (<see cref="AnnotationManager"/>) and the write-side session: an agent that wants
    /// to edit "the indemnification clause" looks the annotation up by id and gets the
    /// anchors it can hand to <see cref="ReplaceText"/> / <see cref="DeleteBlock"/> /
    /// <see cref="Raw"/>. Returns an empty list when the id is unknown or the annotation's
    /// bookmark is missing/malformed.
    /// </summary>
    /// <remarks>
    /// v1 returns the enclosing block anchors — every paragraph/heading/list-item/cell/
    /// row/table whose subtree overlaps the bookmark range. Bookmarks that sit inside a
    /// single paragraph yield that paragraph's anchor; bookmarks spanning multiple blocks
    /// yield each in document order. A finer-grained <see cref="CharSpan"/>-aware return
    /// is left to a follow-up (see the issue's "Out of scope for v1").
    /// </remarks>
    public IReadOnlyList<AnchorTarget> FindByAnnotation(string annotationId)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(annotationId)) return Array.Empty<AnchorTarget>();
        var ann = AnnotationManager.GetAnnotations(_doc!)
            .FirstOrDefault(a => string.Equals(a.Id, annotationId, StringComparison.Ordinal));
        if (ann is null || string.IsNullOrEmpty(ann.BookmarkName))
            return Array.Empty<AnchorTarget>();
        return ResolveBookmarkAnchors(ann.BookmarkName);
    }

    /// <summary>
    /// Finds every annotation whose <see cref="DocumentAnnotation.LabelId"/> equals
    /// <paramref name="labelId"/> and resolves each of their ranges. The result is keyed
    /// by annotation id so callers can disambiguate when the same label was applied to
    /// multiple regions (e.g. three separate "WARRANTY" annotations). Annotations whose
    /// bookmark is missing or resolves to no anchors are omitted from the result.
    /// </summary>
    public IReadOnlyDictionary<string, IReadOnlyList<AnchorTarget>> FindByLabel(string labelId)
    {
        ThrowIfDisposed();
        var map = new Dictionary<string, IReadOnlyList<AnchorTarget>>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(labelId)) return map;
        foreach (var ann in AnnotationManager.GetAnnotations(_doc!))
        {
            if (!string.Equals(ann.LabelId, labelId, StringComparison.Ordinal)) continue;
            if (string.IsNullOrEmpty(ann.BookmarkName)) continue;
            var anchors = ResolveBookmarkAnchors(ann.BookmarkName);
            if (anchors.Count > 0) map[ann.Id] = anchors;
        }
        return map;
    }

    /// <summary>
    /// Resolves any bookmark in the main document part (Docxodus-managed or user-authored)
    /// to the block-level anchors covering its range, in document order. Empty when the
    /// bookmark name is unknown or its end marker is missing. Use this for raw bookmark
    /// names that didn't come from <see cref="AnnotationManager"/>.
    /// </summary>
    public IReadOnlyList<AnchorTarget> FindByBookmark(string bookmarkName)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(bookmarkName)) return Array.Empty<AnchorTarget>();
        return ResolveBookmarkAnchors(bookmarkName);
    }

    /// <summary>
    /// Enumerates every annotation persisted in the document — id, label id/text, color,
    /// author, and (when the bookmark resolves) the annotated text it covers. Lets an
    /// agent prime itself with "here are the labeled regions you can target" before
    /// committing to a specific id.
    /// </summary>
    public IReadOnlyList<DocumentAnnotation> ListAnnotations()
    {
        ThrowIfDisposed();
        return AnnotationManager.GetAnnotations(_doc!);
    }

    /// <summary>
    /// Walks the main document part once: locates the bookmark by name, then collects
    /// every block-level anchor whose subtree overlaps the bookmark range, deduplicated
    /// and sorted in document order. Pre-order positions are recomputed per call rather
    /// than cached — callers in agentic loops should resolve once and reuse the result.
    /// </summary>
    private IReadOnlyList<AnchorTarget> ResolveBookmarkAnchors(string bookmarkName)
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) return Array.Empty<AnchorTarget>();
        var root = main.GetXDocument().Root;
        if (root is null) return Array.Empty<AnchorTarget>();

        var start = root.Descendants(W.bookmarkStart)
            .FirstOrDefault(b => (string?)b.Attribute(W.name) == bookmarkName);
        if (start is null) return Array.Empty<AnchorTarget>();
        var bookmarkId = (string?)start.Attribute(W.id);
        if (bookmarkId is null) return Array.Empty<AnchorTarget>();
        var end = root.Descendants(W.bookmarkEnd)
            .FirstOrDefault(b => (string?)b.Attribute(W.id) == bookmarkId);
        if (end is null) return Array.Empty<AnchorTarget>();

        // Force Project() so Unids are assigned on every block and the AnchorIndex is
        // populated. Building a Unid → AnchorTarget reverse map lets us look up each
        // candidate block without re-running the converter's KindFor classifier here.
        var index = Project().AnchorIndex;
        var byUnid = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
        foreach (var t in index.Values) byUnid[t.Unid] = t;

        // Pre-order positions support two operations: (a) deciding whether a block's
        // subtree overlaps the bookmark range, (b) sorting the collected hits back into
        // document order. O(N) per call — fine for in-session use where Project() is
        // already O(N).
        var pos = new Dictionary<XElement, int>(ReferenceEqualityComparer.Instance);
        int counter = 0;
        foreach (var el in root.DescendantsAndSelf()) pos[el] = counter++;

        if (!pos.TryGetValue(start, out var startPos) || !pos.TryGetValue(end, out var endPos))
            return Array.Empty<AnchorTarget>();
        if (endPos <= startPos) return Array.Empty<AnchorTarget>();

        var hits = new List<(int Pos, AnchorTarget Target)>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var el in root.Descendants())
        {
            var unid = (string?)el.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            if (!byUnid.TryGetValue(unid, out var target)) continue;
            // The bookmark we found lives in the body part, so only body-scope anchors
            // can possibly intersect it. The guard cheaply rejects same-Unid collisions
            // with header/footer/footnote anchors (rare, but possible if the projector's
            // index ever surfaces them).
            if (!string.Equals(target.Anchor.Scope, "body", StringComparison.Ordinal)) continue;

            var elStart = pos[el];
            var lastDesc = el.DescendantsAndSelf().Last();
            var elEnd = pos[lastDesc];
            // Strict overlap on the marker positions themselves: a bookmark sitting
            // exactly between two paragraphs shouldn't pick up either of them.
            if (elEnd <= startPos) continue;
            if (elStart >= endPos) continue;
            if (!seen.Add(target.Anchor.Id)) continue;
            hits.Add((elStart, target));
        }

        hits.Sort((a, b) => a.Pos.CompareTo(b.Pos));
        var result = new AnchorTarget[hits.Count];
        for (int i = 0; i < hits.Count; i++) result[i] = hits[i].Target;
        return result;
    }

    /// <summary>
    /// Surgical text replacement within a single paragraph/heading/list-item: finds every
    /// literal occurrence of <paramref name="find"/> in the anchor's flat text and replaces
    /// it with <paramref name="replace"/>, preserving the surrounding run formatting that
    /// the match didn't touch. Returns one <see cref="EditResult"/> per attempted match.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The replacement text is plain-text and inherits the formatting of the FIRST run the
    /// match spanned — middle/trailing runs keep their <c>w:rPr</c> but lose the slice of
    /// text the match consumed (so a bold run that contributed three chars to the match now
    /// has those three chars gone, but stays bold for everything else it held).
    /// </para>
    /// <para>
    /// Matches are applied in reverse document order so multiple matches in the same
    /// paragraph don't invalidate each other's offsets. The whole call records a single undo
    /// snapshot — <see cref="Undo"/> rolls back every replacement together.
    /// </para>
    /// </remarks>
    public IReadOnlyList<EditResult> ReplaceTextRange(
        string anchorId,
        string find,
        string replace,
        ReplaceOptions? options = null)
    {
        if (_disposed)
            return new[] { EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed") };
        if (string.IsNullOrEmpty(find))
            return new[] { EditResult.Fail(EditErrorCode.MalformedMarkdown, "find must be non-empty", anchorId) };

        var target = FindAnchor(anchorId);
        if (target is null)
            return new[] { EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId) };
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return new[] { EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceTextRange requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId) };

        var opts = options ?? new ReplaceOptions();
        var regexOpts = opts.IgnoreCase
            ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
            : System.Text.RegularExpressions.RegexOptions.None;
        var pattern = System.Text.RegularExpressions.Regex.Escape(find);
        replace = MaybeApplySmartQuotes(replace);

        var matches = Grep(pattern, regexOpts)
            .Where(m => m.EnclosingAnchor.Anchor.Id == target.Anchor.Id)
            .ToList();
        if (opts.MaxReplacements is int cap) matches = matches.Take(cap).ToList();
        if (matches.Count == 0) return Array.Empty<EditResult>();

        var element = target.Resolve(_doc!);
        if (element is null)
            return new[] { EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId) };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Reverse offset order so earlier-offset matches' SpanInElement stays valid
            // after later-offset edits land — see DS112/DS115.
            foreach (var match in matches.OrderByDescending(m => m.Span.Start))
                ApplyFragmentReplacement(element, match, replace);

            InvalidateProjectionCache();
            var success = new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
            return Enumerable.Repeat(success, matches.Count).ToArray();
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return new[] { EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId) };
        }
    }

    /// <summary>
    /// Convenience: replace a single <see cref="TextMatch"/> (typically from <see cref="Grep"/>)
    /// in place with <paramref name="replace"/>. Same fragment-formatting semantics as
    /// <see cref="ReplaceTextRange"/>.
    /// </summary>
    public EditResult ReplaceMatch(TextMatch match, string replace)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");
        return ReplaceTextAtSpan(match.EnclosingAnchor.Anchor.Id, match.Span.Start, match.Span.Length, replace);
    }

    /// <summary>
    /// Replace the bracketed portion of a <see cref="TextMatch"/> with <paramref name="newInner"/>,
    /// preserving any prefix or suffix outside the brackets. Designed for
    /// <see cref="FindPlaceholders"/> matches like <c>$[___]</c> where the regex
    /// <c>\$?\[…\]</c> captures the leading <c>$</c>: <c>ReplaceInner(match, "0.20")</c>
    /// yields <c>$0.20</c> (not <c>0.20</c>). For matches without any prefix/suffix,
    /// this is equivalent to <see cref="ReplaceMatch"/> with the new inner value.
    /// Returns <see cref="EditErrorCode.MalformedMarkdown"/> if the match text does
    /// not contain balanced brackets.
    /// </summary>
    public EditResult ReplaceInner(TextMatch match, string newInner)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");

        int lb = match.Text.IndexOf('[');
        int rb = match.Text.LastIndexOf(']');
        if (lb < 0 || rb <= lb)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown,
                $"match text has no balanced brackets: '{match.Text}'");

        var prefix = match.Text[..lb];
        var suffix = match.Text[(rb + 1)..];
        return ReplaceMatch(match, prefix + newInner + suffix);
    }

    /// <summary>
    /// Surgical replacement of an exact byte range within one block's flat text.
    /// The natural pair to <see cref="Grep"/>: pass the <see cref="TextMatch.EnclosingAnchor"/>'s
    /// id plus the <see cref="TextMatch.Span"/> coordinates to replace one specific match
    /// even when several identical needles share the same paragraph (the template-filling
    /// case where five <c>[___]</c> placeholders each get a different value).
    /// </summary>
    public EditResult ReplaceTextAtSpan(string anchorId, int spanStart, int spanLength, string replace)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceTextAtSpan requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        replace = MaybeApplySmartQuotes(replace);

        var map = Internal.RunTextMap.Build(element);
        if (spanStart < 0 || spanLength < 0 || spanStart + spanLength > map.FlatText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span {spanStart}+{spanLength} out of [0, {map.FlatText.Length}]", anchorId);

        var pieces = Internal.RunTextMap.ResolveRange(map, spanStart, spanLength);
        if (pieces.Count == 0)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange, "span resolved to no runs", anchorId);

        // Synthesize fragments from the resolved pieces. The replacement helper only
        // reads Unid + SpanInElement, so the other fields are placeholders.
        var fragments = new List<RunFragment>(pieces.Count);
        foreach (var (seg, offsetInRun, len) in pieces)
        {
            var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
            fragments.Add(new RunFragment
            {
                Unid = runUnid,
                Text = string.Empty,
                SpanInElement = new CharSpan(offsetInRun, len),
                Formatting = new RunFormatting(),
            });
        }
        var synthetic = new TextMatch
        {
            Text = map.FlatText.Substring(spanStart, spanLength),
            EnclosingAnchor = target,
            Span = new CharSpan(spanStart, spanLength),
            Fragments = fragments,
            ContextBefore = string.Empty,
            ContextAfter = string.Empty,
        };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            ApplyFragmentReplacement(element, synthetic, replace);
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Enumerate the template placeholders in the document. A thin classifier over
    /// <see cref="Grep"/> that distinguishes <c>[___]</c> value blanks, <c>[bracketed
    /// alternative clauses]</c>, and <c>[insert X]</c> / <c>[*italic hint*]</c>
    /// instruction placeholders — the three families a template-filling agent treats
    /// differently. See <see cref="PlaceholderKind"/> for the taxonomy.
    /// </summary>
    /// <remarks>
    /// Nested brackets resolve to the INNERMOST bracket. A construct like
    /// <c>[under the name [Bluth Co.]]</c> produces a placeholder for the inner
    /// <c>[Bluth Co.]</c> only — usually what an agent cares about — but the outer
    /// optional-clause bracket isn't reported separately. Use <see cref="Grep"/> with
    /// a balanced-bracket regex if you need both.
    /// </remarks>
    public IReadOnlyList<TemplatePlaceholder> FindPlaceholders(
        PlaceholderKinds kinds = PlaceholderKinds.All,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 80,
        ContextBoundary boundary = ContextBoundary.Char)
    {
        ThrowIfDisposed();
        if (kinds == 0) return Array.Empty<TemplatePlaceholder>();

        // Single bracket-or-dollar-bracket scan; classify by content after the match.
        // Non-greedy inner content + negated character class keeps the regex from
        // crossing into a sibling bracket pair on the same line.
        var matches = Grep(@"\$?\[[^\[\]]+\]",
            System.Text.RegularExpressions.RegexOptions.None, scope,
            contextChars, WhitespaceMode.Preserve, boundary);
        var results = new List<TemplatePlaceholder>(matches.Count);
        foreach (var m in matches)
        {
            var (classified, alternatives) = Classify(m.Text);
            if (classified is not PlaceholderKind kind) continue;
            if (!kinds.HasFlag(KindToFlag(kind))) continue;
            results.Add(new TemplatePlaceholder
            {
                Match = m,
                Kind = kind,
                Hint = kind == PlaceholderKind.Instruction ? ExtractHint(m.Text) : null,
                AlternativeKinds = alternatives,
            });
        }
        return results;

        static (PlaceholderKind? Primary, IReadOnlyList<PlaceholderKind> Alternatives) Classify(string text)
        {
            var inner = text.StartsWith('$') ? text[2..^1] : text[1..^1];

            // BlankFill: 2+ underscores anywhere inside (so "[__]" director-count slots,
            // "[___ times]" unit-suffix slots, and "[________ __, 20__]" date-shaped
            // slots all qualify). Tighter than "any underscore" to avoid false positives
            // on quoted identifiers like "[a_b]". Trade-off in writeup at the FindPlaceholders
            // section of docs/architecture/docx_mutation_api.md.
            bool isBlankFill = inner.Count(c => c == '_') >= 2;

            // Instruction: italicized (asterisk-wrapped) text, or starts with the
            // drafter verbs "insert" / "specify". Conservative leading-word check
            // so general prose in brackets doesn't mis-classify.
            bool isInstruction = false;
            if (inner.StartsWith('*') && inner.EndsWith('*') && inner.Length > 2) isInstruction = true;
            else
            {
                var firstWord = inner.TakeWhile(char.IsLetter).ToArray();
                var w = new string(firstWord).ToLowerInvariant();
                if (w is "insert" or "specify") isInstruction = true;
            }

            // Secondary classification: long-clause-with-blanks. When BlankFill fires but
            // the inner text reads like a multi-word clause (4+ spaces between words),
            // the placeholder is plausibly an AlternativeClause with an embedded blank.
            // Caller can detect via AlternativeKinds and strip the outer brackets, then
            // separately fill the inner _______ run.
            bool looksClause = inner.Count(c => c == ' ') >= 4;

            // Primary classification keeps the original priority order:
            //   BlankFill → Instruction → AlternativeClause
            if (isBlankFill)
            {
                var alts = looksClause ? new[] { PlaceholderKind.AlternativeClause } : Array.Empty<PlaceholderKind>();
                return (PlaceholderKind.BlankFill, alts);
            }
            if (isInstruction)
                return (PlaceholderKind.Instruction, Array.Empty<PlaceholderKind>());
            return (PlaceholderKind.AlternativeClause, Array.Empty<PlaceholderKind>());
        }

        static string ExtractHint(string text)
        {
            var inner = text.StartsWith('$') ? text[2..^1] : text[1..^1];
            // Strip a single pair of surrounding asterisks (italic markers from the projector).
            if (inner.StartsWith('*') && inner.EndsWith('*') && inner.Length > 2)
                inner = inner[1..^1];
            return inner.Trim();
        }

        static PlaceholderKinds KindToFlag(PlaceholderKind k) => k switch
        {
            PlaceholderKind.BlankFill => PlaceholderKinds.BlankFill,
            PlaceholderKind.AlternativeClause => PlaceholderKinds.AlternativeClause,
            PlaceholderKind.Instruction => PlaceholderKinds.Instruction,
            _ => 0,
        };
    }

    /// <summary>
    /// Compose a high-signal snapshot of the session's edit-state — total anchors,
    /// remaining bracketed placeholders, bare underscore runs, and footnote/comment
    /// counts. Pure composition of existing primitives (<see cref="Project"/>,
    /// <see cref="FindPlaceholders"/>, <see cref="Grep"/>) with no new logic, so
    /// every count is exactly what the caller would compute by hand. Designed as
    /// the canonical "what's left to fill in?" check after a mutation batch.
    /// </summary>
    /// <remarks>
    /// The bare-underscore regex <c>(?&lt;![\[_])_{3,}(?![\]_])</c> uses lookarounds
    /// that exclude both a bracket and an adjacent underscore, so they guard the
    /// boundaries of the maximal underscore run (not just the regex match) and
    /// avoid false positives inside <c>[_____]</c>. Bracketed underscore runs are
    /// surfaced via <see cref="EditSummary.RemainingPlaceholders"/>, so the two
    /// collections are disjoint by construction. Both queries run against
    /// <see cref="ProjectionScopes.All"/> so headers/footers/footnotes/endnotes/comments
    /// are counted symmetrically.
    /// </remarks>
    public EditSummary GetEditSummary()
    {
        ThrowIfDisposed();

        var projection = Project();
        var placeholders = FindPlaceholders(PlaceholderKinds.All, ProjectionScopes.All);
        var underscoreRuns = Grep(@"(?<![\[_])_{3,}(?![\]_])", scope: ProjectionScopes.All);

        int footnoteCount = 0;
        int commentCount = 0;
        foreach (var t in projection.AnchorIndex.Values)
        {
            if (t.Anchor.Kind == "fn" && t.Anchor.Scope == "fn") footnoteCount++;
            else if (t.Anchor.Kind == "cmt" && t.Anchor.Scope == "cmt") commentCount++;
        }

        var main = _doc!.MainDocumentPart;
        int inlineFnRefs = 0;
        if (main is not null)
            inlineFnRefs = main.GetXDocument().Root!.Descendants(W.footnoteReference).Count();

        return new EditSummary
        {
            TotalAnchors = projection.AnchorIndex.Count,
            RemainingPlaceholders = placeholders,
            BareUnderscoreRuns = underscoreRuns,
            FootnoteCount = footnoteCount,
            InlineFootnoteRefCount = inlineFnRefs,
            CommentCount = commentCount,
        };
    }

    /// <summary>
    /// Thin discoverability alias for <see cref="FindPlaceholders"/>. Same return
    /// shape; the rename exists because "what's remaining?" reads more naturally
    /// at agent call sites than "find the placeholders."
    /// </summary>
    public IReadOnlyList<TemplatePlaceholder> RemainingPlaceholders(
        PlaceholderKinds kinds = PlaceholderKinds.All) =>
        FindPlaceholders(kinds);

    /// <summary>
    /// Diffs the projection captured at session construction against the current projection
    /// and returns an anchor-keyed change list. Keyed by <c>(scope, Unid)</c> — the Unid
    /// is stable across mutations and kind flips (a paragraph promoted to a heading keeps
    /// its Unid while its anchor kind goes from "p" to "h"), and the scope qualifier guards
    /// against cross-part Unid collisions (the deterministic Unid scheme seeds each scope's
    /// root with the root element's local name, so two header parts whose first paragraph
    /// has identical structure end up with the same raw Unid in different scopes — see
    /// issue #187). Requires <see cref="DocxSessionSettings.CaptureInitialProjection"/>
    /// to have been <c>true</c> at construction time.
    /// </summary>
    /// <param name="format">Output shape. <see cref="DiffFormat.Json"/> (default) returns
    /// an anchor-keyed JSON array; <see cref="DiffFormat.Unified"/> returns a
    /// <c>patch(1)</c>-compatible unified diff over the markdown projections;
    /// <see cref="DiffFormat.SideBySide"/> returns a two-column human-review diff.</param>
    /// <returns>For <see cref="DiffFormat.Json"/>, a JSON array of <see cref="DiffEntry"/>
    /// records. Entries are grouped by op (all deletes first, then modifies, then inserts);
    /// within each group, by anchor-index iteration order (which is document order in
    /// practice, since the projector builds the index via a depth-first descendant walk).
    /// Returns <c>"[]"</c> when the document has not been mutated since construction.
    /// For <see cref="DiffFormat.Unified"/>, a standard unified diff with <c>--- initial</c>
    /// / <c>+++ current</c> headers and 3 lines of context; empty string when nothing changed.
    /// For <see cref="DiffFormat.SideBySide"/>, a two-column rendering with the initial
    /// projection padded to 72 chars on the left, a single marker character, then the
    /// current projection.</returns>
    /// <exception cref="InvalidOperationException">Thrown when
    /// <see cref="DocxSessionSettings.CaptureInitialProjection"/> was <c>false</c>.</exception>
    /// <exception cref="NotSupportedException">Thrown for <paramref name="format"/> values
    /// outside the defined <see cref="DiffFormat"/> range.</exception>
    public string GetDiff(DiffFormat format = DiffFormat.Json)
    {
        ThrowIfDisposed();
        if (_initialProjection is null)
            throw new InvalidOperationException(
                "GetDiff requires CaptureInitialProjection = true in DocxSessionSettings.");

        var current = Project();

        return format switch
        {
            DiffFormat.Json => SerializeDiff(ComputeDiff(_initialProjection, current)),
            DiffFormat.Unified => SerializeUnifiedDiff(_initialProjection.Markdown, current.Markdown),
            DiffFormat.SideBySide => SerializeSideBySideDiff(_initialProjection.Markdown, current.Markdown),
            _ => throw new NotSupportedException(
                $"DiffFormat.{format} is not a recognized value."),
        };
    }

    private static List<DiffEntry> ComputeDiff(MarkdownProjection initial, MarkdownProjection current)
    {
        // Key by (scope, Unid). Two reasons we cannot use Unid alone:
        //   1. AnchorIndex is dual-keyed under non-FullUnid rendering (the same
        //      AnchorTarget is reachable via its full Unid key and its rendered
        //      alias key), so AnchorIndex.Values enumerates the same target twice.
        //   2. The deterministic Unid scheme seeds each scope's root with the root
        //      element's local name ("hdr" for every header part, "ftr" for every
        //      footer part), so two header parts whose first paragraph has the
        //      same content + position end up with identical raw Unids in
        //      different scopes (reproduced on the NVCA Model COI — issue #187).
        // DistinctBy collapses duplicates from case (1); the composite key
        // separates legitimately distinct targets from case (2).
        var initialByKey = initial.AnchorIndex.Values
            .DistinctBy(t => (t.Anchor.Scope, t.Unid))
            .ToDictionary(t => (t.Anchor.Scope, t.Unid));
        var currentByKey = current.AnchorIndex.Values
            .DistinctBy(t => (t.Anchor.Scope, t.Unid))
            .ToDictionary(t => (t.Anchor.Scope, t.Unid));

        var entries = new List<DiffEntry>();

        // Deletes: in initial, missing from current.
        foreach (var (key, target) in initialByKey)
        {
            if (currentByKey.ContainsKey(key)) continue;
            entries.Add(new DiffEntry
            {
                Op = "delete",
                AnchorId = target.Anchor.Id,
                Before = target.TextPreview,
            });
        }

        // Modifies: present in both, text preview OR kind differs.
        // Kind can flip without a text change (e.g., SetParagraphStyle promoting
        // a paragraph to a heading flips Anchor.Kind from "p" to "h" while
        // preserving the Unid and TextPreview).
        foreach (var (key, initialTarget) in initialByKey)
        {
            if (!currentByKey.TryGetValue(key, out var currentTarget)) continue;
            if (initialTarget.TextPreview == currentTarget.TextPreview
                && initialTarget.Anchor.Kind == currentTarget.Anchor.Kind) continue;
            entries.Add(new DiffEntry
            {
                Op = "modify",
                AnchorId = currentTarget.Anchor.Id,
                Before = initialTarget.TextPreview,
                After = currentTarget.TextPreview,
            });
        }

        // Inserts: in current, missing from initial.
        foreach (var (key, target) in currentByKey)
        {
            if (initialByKey.ContainsKey(key)) continue;
            entries.Add(new DiffEntry
            {
                Op = "insert",
                AnchorId = target.Anchor.Id,
                After = target.TextPreview,
            });
        }

        return entries;
    }

    private static string SerializeDiff(List<DiffEntry> entries)
    {
        // Hand-rolled JSON so SerializeDiff stays trim/AOT-safe; the WASM build
        // ships with reflection-based serialization disabled, so
        // `System.Text.Json.JsonSerializer.Serialize(...)` throws
        // `JsonSerializerIsReflectionDisabled` at runtime in the browser.
        if (entries.Count == 0) return "[]";
        var sb = new System.Text.StringBuilder(entries.Count * 100 + 2);
        sb.Append('[');
        for (int i = 0; i < entries.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var e = entries[i];
            sb.Append("{\"op\":\"").Append(e.Op).Append("\"")
              .Append(",\"anchorId\":");
            AppendJsonString(sb, e.AnchorId);
            if (e.Before is not null)
            {
                sb.Append(",\"before\":");
                AppendJsonString(sb, e.Before);
            }
            if (e.After is not null)
            {
                sb.Append(",\"after\":");
                AppendJsonString(sb, e.After);
            }
            sb.Append('}');
        }
        sb.Append(']');
        return sb.ToString();
    }

    private static void AppendJsonString(System.Text.StringBuilder sb, string s)
    {
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("X4"));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
    }

    // ─── Line-based LCS for DiffFormat.Unified / SideBySide ────────────────
    //
    // Hand-rolled O(n*m) LCS over arrays of lines. We deliberately avoid pulling
    // in DiffPlex / DiffMatchPatch — the WASM build disables reflection-based
    // serialization and we want this path to stay AOT-friendly without a NuGet
    // edge case. The unified path is parseable by patch(1); the side-by-side
    // path mirrors `diff -y` markers.

    private enum LineDiffKind { Equal, Delete, Insert }

    private readonly record struct LineDiffOp(LineDiffKind Kind, int AIdx, int BIdx);

    private static List<LineDiffOp> ComputeLineDiff(string[] a, string[] b)
    {
        int n = a.Length, m = b.Length;
        // dp[i, j] = length of LCS of a[..i] and b[..j].
        var dp = new int[n + 1, m + 1];
        for (int i = 1; i <= n; i++)
        {
            for (int j = 1; j <= m; j++)
            {
                dp[i, j] = a[i - 1] == b[j - 1]
                    ? dp[i - 1, j - 1] + 1
                    : Math.Max(dp[i - 1, j], dp[i, j - 1]);
            }
        }

        var ops = new List<LineDiffOp>(n + m);
        int x = n, y = m;
        while (x > 0 && y > 0)
        {
            if (a[x - 1] == b[y - 1])
            {
                ops.Add(new LineDiffOp(LineDiffKind.Equal, x - 1, y - 1));
                x--; y--;
            }
            else if (dp[x - 1, y] > dp[x, y - 1])
            {
                ops.Add(new LineDiffOp(LineDiffKind.Delete, x - 1, -1));
                x--;
            }
            else
            {
                // Ties (dp[x-1,y] == dp[x,y-1]) go to Insert during backward traversal
                // so that after List.Reverse() the forward order shows Delete before
                // Insert — the conventional ordering for unified diffs and the
                // precondition for `SerializeSideBySideDiff`'s Delete+Insert →
                // "modify" pairing.
                ops.Add(new LineDiffOp(LineDiffKind.Insert, -1, y - 1));
                y--;
            }
        }
        while (x > 0) { ops.Add(new LineDiffOp(LineDiffKind.Delete, x - 1, -1)); x--; }
        while (y > 0) { ops.Add(new LineDiffOp(LineDiffKind.Insert, -1, y - 1)); y--; }

        ops.Reverse();
        return ops;
    }

    private const int UnifiedContextLines = 3;

    private static string SerializeUnifiedDiff(string initial, string current)
    {
        // Split on '\n' only — the markdown projector emits LF line terminators.
        // Trailing '\n' produces a trailing empty element; that round-trips
        // correctly through patch(1) provided we don't add a phantom newline.
        var a = initial.Split('\n');
        var b = current.Split('\n');
        var ops = ComputeLineDiff(a, b);

        // No changes → empty string. Lets `if (string.IsNullOrEmpty(diff))` be the
        // "did anything change?" check on the call site.
        bool anyChange = false;
        for (int i = 0; i < ops.Count; i++)
        {
            if (ops[i].Kind != LineDiffKind.Equal) { anyChange = true; break; }
        }
        if (!anyChange) return string.Empty;

        var sb = new System.Text.StringBuilder();
        sb.Append("--- initial\n");
        sb.Append("+++ current\n");

        int idx = 0;
        while (idx < ops.Count)
        {
            // Skip leading Equal ops between hunks.
            while (idx < ops.Count && ops[idx].Kind == LineDiffKind.Equal) idx++;
            if (idx >= ops.Count) break;

            int hunkStart = Math.Max(0, idx - UnifiedContextLines);
            int lastChange = idx;
            int scan = idx;
            while (scan < ops.Count)
            {
                if (ops[scan].Kind != LineDiffKind.Equal)
                {
                    lastChange = scan;
                    scan++;
                    continue;
                }

                // Break when we'd have more than 2 * contextLines equal ops between
                // the last change and the next one — that's where one hunk ends and
                // the next begins.
                int gap = 0;
                while (scan < ops.Count && ops[scan].Kind == LineDiffKind.Equal)
                {
                    gap++;
                    if (gap > 2 * UnifiedContextLines) break;
                    scan++;
                }
                if (gap > 2 * UnifiedContextLines) break;
            }
            int hunkEnd = Math.Min(ops.Count, lastChange + UnifiedContextLines + 1);

            // Compute 1-based line numbers and counts for the hunk header.
            int aStart = 0, bStart = 0;
            for (int k = 0; k < hunkStart; k++)
            {
                if (ops[k].Kind != LineDiffKind.Insert) aStart++;
                if (ops[k].Kind != LineDiffKind.Delete) bStart++;
            }
            int aLines = 0, bLines = 0;
            for (int k = hunkStart; k < hunkEnd; k++)
            {
                if (ops[k].Kind != LineDiffKind.Insert) aLines++;
                if (ops[k].Kind != LineDiffKind.Delete) bLines++;
            }

            // Unified-diff convention: when count is 0, the start position is the
            // line *before* the change (so a pure-insert hunk reads "@@ -0,0 +1,N @@").
            // When count is >0, we emit "start+1" to convert from 0-based to 1-based.
            int aHeaderStart = aLines == 0 ? aStart : aStart + 1;
            int bHeaderStart = bLines == 0 ? bStart : bStart + 1;

            sb.Append("@@ -").Append(aHeaderStart).Append(',').Append(aLines)
              .Append(" +").Append(bHeaderStart).Append(',').Append(bLines)
              .Append(" @@\n");

            for (int k = hunkStart; k < hunkEnd; k++)
            {
                var op = ops[k];
                switch (op.Kind)
                {
                    case LineDiffKind.Equal:
                        sb.Append(' ').Append(a[op.AIdx]).Append('\n');
                        break;
                    case LineDiffKind.Delete:
                        sb.Append('-').Append(a[op.AIdx]).Append('\n');
                        break;
                    case LineDiffKind.Insert:
                        sb.Append('+').Append(b[op.BIdx]).Append('\n');
                        break;
                }
            }

            idx = hunkEnd;
        }

        return sb.ToString();
    }

    private const int SideBySideColumnWidth = 72;

    private static string SerializeSideBySideDiff(string initial, string current)
    {
        var a = initial.Split('\n');
        var b = current.Split('\n');
        var ops = ComputeLineDiff(a, b);
        if (ops.Count == 0) return string.Empty;

        var sb = new System.Text.StringBuilder();
        int i = 0;
        while (i < ops.Count)
        {
            var op = ops[i];

            // Pair an adjacent Delete + Insert into a single "modified" row marked
            // '|' — matches `diff -y`'s presentation and keeps the row count tight
            // when text on a line is rewritten in place.
            if (op.Kind == LineDiffKind.Delete
                && i + 1 < ops.Count
                && ops[i + 1].Kind == LineDiffKind.Insert)
            {
                AppendSideBySideRow(sb, a[op.AIdx], b[ops[i + 1].BIdx], '|');
                i += 2;
                continue;
            }

            switch (op.Kind)
            {
                case LineDiffKind.Equal:
                    AppendSideBySideRow(sb, a[op.AIdx], b[op.BIdx], ' ');
                    break;
                case LineDiffKind.Delete:
                    AppendSideBySideRow(sb, a[op.AIdx], string.Empty, '<');
                    break;
                case LineDiffKind.Insert:
                    AppendSideBySideRow(sb, string.Empty, b[op.BIdx], '>');
                    break;
            }
            i++;
        }

        return sb.ToString();
    }

    private static void AppendSideBySideRow(System.Text.StringBuilder sb, string left, string right, char marker)
    {
        // Truncate (with U+2026 tail) anything past the column width so the marker
        // column stays aligned. The right column is allowed to run to end-of-line —
        // a terminal will wrap it; a viewer that hard-wraps can post-process.
        string leftDisp = left.Length > SideBySideColumnWidth
            ? string.Concat(left.AsSpan(0, SideBySideColumnWidth - 1), "…")
            : left.PadRight(SideBySideColumnWidth);
        sb.Append(leftDisp).Append(' ').Append(marker).Append(' ').Append(right).Append('\n');
    }

    /// <summary>
    /// Picker-driven template fill. For every placeholder matching
    /// <see cref="FillOptions.Kinds"/>, calls <paramref name="picker"/>; if the picker
    /// returns a non-null string, the placeholder is replaced (with optional
    /// <c>$</c>-prefix preservation per <see cref="FillOptions.PreserveDollarPrefix"/>).
    /// Iterates until no more placeholders match (or until <see cref="FillOptions.MaxPasses"/>
    /// is reached, or a pass makes zero state changes) — important when
    /// <see cref="FillOptions.Kinds"/> includes <see cref="PlaceholderKinds.AlternativeClause"/>
    /// and the doc has nested brackets that surface only after the inner ones are stripped.
    /// Replacements within a paragraph are applied in reverse-offset order automatically.
    /// The picker may be invoked more than once for the same logical placeholder
    /// when <see cref="FillOptions.Kinds"/> includes <see cref="PlaceholderKinds.AlternativeClause"/>
    /// and inner brackets are stripped between passes; pickers must therefore be
    /// deterministic on <c>p.Match.Text</c> (return the same result for the same
    /// input text). Non-deterministic pickers can produce inconsistent fills.
    /// </summary>
    public BulkEditResult FillPlaceholders(
        Func<TemplatePlaceholder, string?> picker,
        FillOptions? options = null)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(picker);
        var opts = options ?? new FillOptions();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(opts.MaxPasses);

        int filled = 0;
        int workPasses = 0;
        var errors = new List<EditError>();
        var unfilled = new List<TemplatePlaceholder>();
        var seenSkipKeys = new HashSet<(string AnchorId, int Start, int Length)>();

        for (int pass = 1; pass <= opts.MaxPasses; pass++)
        {
            var placeholders = FindPlaceholders(opts.Kinds, opts.Scope, opts.ContextChars, opts.Boundary)
                .OrderByDescending(p => p.Match.EnclosingAnchor.Anchor.Id, StringComparer.Ordinal)
                .ThenByDescending(p => p.Match.Span.Start)
                .ToList();
            if (placeholders.Count == 0) break;

            int passChanges = 0;
            foreach (var p in placeholders)
            {
                var pick = picker(p);
                if (pick is null)
                {
                    // Count each skip exactly once per placeholder lifetime.
                    var key = (p.Match.EnclosingAnchor.Anchor.Id, p.Match.Span.Start, p.Match.Span.Length);
                    if (seenSkipKeys.Add(key))
                        unfilled.Add(p);
                    continue;
                }

                if (opts.PreserveDollarPrefix && p.Match.Text.StartsWith("$") && !pick.StartsWith("$"))
                    pick = "$" + pick;

                var r = opts.CoalesceWhitespaceAroundEmptyFill && pick.Length == 0
                    ? ReplaceMatchCoalescingNeighbors(p.Match)
                    : ReplaceMatch(p.Match, pick);
                if (r.Success)
                {
                    filled++;
                    passChanges++;
                }
                else if (r.Error is { } err)
                {
                    errors.Add(err);
                }
            }

            // Record this pass only if it did real work — observation alone
            // (placeholders found but all skipped or all errored) doesn't count.
            if (passChanges > 0)
                workPasses = pass;

            // If this pass made no changes, the picker is steady-state — stop iterating.
            if (passChanges == 0) break;
        }

        int stillPresent = FindPlaceholders(opts.Kinds, opts.Scope).Count;

        return new BulkEditResult
        {
            Filled = filled,
            Skipped = unfilled.Count,
            StillPresent = stillPresent,
            Passes = workPasses,
            Unfilled = unfilled,
            Errors = errors,
        };
    }

    /// <summary>
    /// Helper for <see cref="FillPlaceholders"/>'s
    /// <see cref="FillOptions.CoalesceWhitespaceAroundEmptyFill"/> path: deletes the
    /// match's span and, based on the chars immediately adjacent in the enclosing
    /// block's flat text, also absorbs surrounding whitespace / leading-space-before-punctuation
    /// / matched-brackets. See the option's docs for the exact rules. Falls back
    /// to a literal <see cref="ReplaceMatch"/> with empty string when no neighbor
    /// pattern matches.
    /// </summary>
    private EditResult ReplaceMatchCoalescingNeighbors(TextMatch match)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var anchorId = match.EnclosingAnchor.Anchor.Id;
        var target = FindAnchor(anchorId);
        if (target is null) return ReplaceMatch(match, string.Empty);
        var element = target.Resolve(_doc!);
        if (element is null) return ReplaceMatch(match, string.Empty);

        var flat = Internal.RunTextMap.Build(element).FlatText;
        int start = match.Span.Start;
        int end = start + match.Span.Length;
        if (start < 0 || end > flat.Length) return ReplaceMatch(match, string.Empty);

        char? leftChar = start > 0 ? flat[start - 1] : null;
        char? rightChar = end < flat.Length ? flat[end] : null;

        // Fold the Unicode whitespace variants Word documents commonly use
        // (NBSP, narrow NBSP, thin space) to ASCII space for the rules below so
        // an NBSP-on-either-side still gets coalesced like a regular space.
        static char? Fold(char? c) => c switch
        {
            ' ' or ' ' or ' ' => ' ',
            _ => c,
        };
        char? l = Fold(leftChar);
        char? r = Fold(rightChar);

        static bool IsAsciiSpace(char? c) => c is ' ' or '\t';
        static bool IsClauseTerminator(char? c) => c is '.' or ',' or ';' or ':' or '!' or '?';
        static bool IsOpenBracket(char? c) => c is '(' or '[' or '{';
        static bool IsCloseBracket(char? c) => c is ')' or ']' or '}';

        int extendLeft = 0;
        int extendRight = 0;

        if (IsAsciiSpace(l) && IsAsciiSpace(r))
        {
            // " [x] " → consume the trailing space, leaving one space.
            extendRight = 1;
        }
        else if (IsAsciiSpace(l) && IsClauseTerminator(r))
        {
            // " [x]." / " [x]," → drop the leading space.
            extendLeft = 1;
        }
        else if (IsOpenBracket(l) && IsCloseBracket(r))
        {
            // "([x])" / "[[x]]" → drop both surrounding brackets.
            extendLeft = 1;
            extendRight = 1;
        }

        if (extendLeft == 0 && extendRight == 0)
            return ReplaceMatch(match, string.Empty);

        return ReplaceTextAtSpan(
            anchorId,
            start - extendLeft,
            match.Span.Length + extendLeft + extendRight,
            string.Empty);
    }

    /// <summary>
    /// Apply <paramref name="match"/>'s fragment list to the live element, inserting
    /// <paramref name="replace"/> into the first fragment's run and removing each
    /// subsequent fragment's slice from its run (preserving each run's rPr).
    /// </summary>
    private static void ApplyFragmentReplacement(XElement blockElement, TextMatch match, string replace)
    {
        if (match.Fragments.Count == 0) return;

        // Build a unid → XElement run lookup once. The run XElements are the live
        // descendants of `blockElement` (walking hyperlink/sdt containers too).
        var runsByUnid = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var run in InlineRuns(blockElement))
        {
            var unid = (string?)run.Attribute(PtOpenXml.Unid);
            if (unid is not null) runsByUnid[unid] = run;
        }

        for (int i = 0; i < match.Fragments.Count; i++)
        {
            var fragment = match.Fragments[i];
            if (!runsByUnid.TryGetValue(fragment.Unid, out var run)) continue;

            var concat = RunText(run);
            var start = fragment.SpanInElement.Start;
            var len = fragment.SpanInElement.Length;
            if (start < 0 || start + len > concat.Length) continue;

            var before = concat.Substring(0, start);
            var after = concat.Substring(start + len);
            var newText = i == 0 ? before + replace + after : before + after;

            // Collapse all w:t descendants in this run into a single w:t with the new text.
            // Loses any inline <w:tab/>/<w:br/> inside the run's text section — they're rare
            // for placeholder slots and supporting them here would balloon the impl. Run's
            // rPr/proofErr siblings are untouched, which is the formatting-preservation contract.
            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                newText));
        }
    }

    /// <summary>
    /// When <see cref="DocxSessionSettings.SmartQuotes"/> is on, replace ASCII <c>"</c>
    /// and <c>'</c> with typographic curly quotes. Heuristic: open quote at the start
    /// of the string, after whitespace, or after an open-bracket-like character;
    /// close quote everywhere else. 1:1 character substitution preserves offsets so
    /// downstream span math stays correct.
    /// </summary>
    private string MaybeApplySmartQuotes(string text)
    {
        if (!_settings.SmartQuotes || string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c != '"' && c != '\'') { sb.Append(c); continue; }

            // Look at the previous character (default to "start of string" = whitespace).
            char prev = i == 0 ? ' ' : text[i - 1];
            bool open = char.IsWhiteSpace(prev) || prev is '(' or '[' or '{' or '<';

            sb.Append(c switch
            {
                '"' => open ? '“' : '”',
                '\'' => open ? '‘' : '’',
                _ => c,
            });
        }
        return sb.ToString();
    }

    /// <summary>
    /// Maps the Unicode whitespace variants Word documents commonly use (NBSP, narrow
    /// NBSP, thin space) to ASCII space. Each substitution is one-character-for-one,
    /// so character offsets in the result map 1:1 to the input.
    /// </summary>
    private static string NormalizeWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new System.Text.StringBuilder(text.Length);
        foreach (var c in text)
        {
            sb.Append(c switch
            {
                ' ' => ' ', // non-breaking space
                ' ' => ' ', // narrow no-break space
                ' ' => ' ', // thin space
                _ => c,
            });
        }
        return sb.ToString();
    }

    /// <summary>
    /// Walks outward from a match span by character, stopping at either the
    /// <c>contextChars</c> cap or the nearest character that qualifies as a
    /// boundary under <paramref name="boundary"/>. Returns the <c>(before, after)</c>
    /// text slices. Used by both <see cref="Grep"/> and <see cref="GrepCrossBlock"/>.
    /// </summary>
    private static (string Before, string After) WalkContext(
        string text, int matchStart, int matchLength, int contextChars, ContextBoundary boundary)
    {
        int matchEnd = matchStart + matchLength;

        int leftCap = Math.Max(0, matchStart - contextChars);
        int leftStop = matchStart;
        while (leftStop > leftCap)
        {
            if (IsBoundary(text[leftStop - 1], boundary)) break;
            leftStop--;
        }

        int rightCap = Math.Min(text.Length, matchEnd + contextChars);
        int rightStop = matchEnd;
        while (rightStop < rightCap)
        {
            if (IsBoundary(text[rightStop], boundary)) break;
            rightStop++;
        }

        return (text.Substring(leftStop, matchStart - leftStop),
                text.Substring(matchEnd, rightStop - matchEnd));
    }

    private static bool IsBoundary(char c, ContextBoundary mode) => mode switch
    {
        ContextBoundary.Char => false,
        ContextBoundary.Bracket => c is '[' or ']',
        ContextBoundary.Sentence => c is '.' or '!' or '?' or ':' or ';',
        ContextBoundary.Comma => c is ',',
        _ => false,
    };

    private static bool ScopeMatches(string anchorScope, ProjectionScopes filter)
    {
        // Anchor scopes are strings ("body", "hdr1", "ftr2", "fn", "en", "cmt").
        // ProjectionScopes is a flags enum over the same categories.
        if (anchorScope == "body") return filter.HasFlag(ProjectionScopes.Body);
        if (anchorScope.StartsWith("hdr", StringComparison.Ordinal)) return filter.HasFlag(ProjectionScopes.Headers);
        if (anchorScope.StartsWith("ftr", StringComparison.Ordinal)) return filter.HasFlag(ProjectionScopes.Footers);
        if (anchorScope == "fn") return filter.HasFlag(ProjectionScopes.Footnotes);
        if (anchorScope == "en") return filter.HasFlag(ProjectionScopes.Endnotes);
        if (anchorScope == "cmt") return filter.HasFlag(ProjectionScopes.Comments);
        return false;
    }

    private OpenXmlPart? ResolvePart(string partUri) =>
        EnumerateProjectedParts().FirstOrDefault(p => p.Uri.ToString() == partUri);

    private static RunFormatting ExtractFormatting(XElement run, OpenXmlPart? ownerPart)
    {
        var rPr = run.Element(W.rPr);
        string? hyperlinkUrl = null;
        for (var p = run.Parent; p is not null; p = p.Parent)
        {
            if (p.Name == W.hyperlink)
            {
                var rid = (string?)p.Attribute(R.id);
                if (!string.IsNullOrEmpty(rid) && ownerPart is not null)
                {
                    var rel = ownerPart.HyperlinkRelationships.FirstOrDefault(x => x.Id == rid);
                    if (rel is not null) hyperlinkUrl = rel.Uri.ToString();
                }
                break;
            }
        }

        return new RunFormatting
        {
            Bold = rPr?.Element(W.b) is not null,
            Italic = rPr?.Element(W.i) is not null,
            Underline = rPr?.Element(W.u) is not null,
            Strike = rPr?.Element(W.strike) is not null,
            Code = string.Equals((string?)rPr?.Element(W.rStyle)?.Attribute(W.val), "Code", StringComparison.Ordinal),
            Color = (string?)rPr?.Element(W.color)?.Attribute(W.val),
            HyperlinkUrl = hyperlinkUrl,
            RunStyle = (string?)rPr?.Element(W.rStyle)?.Attribute(W.val),
        };
    }

    /// <summary>
    /// Serialize the current document state. Anchor-id bookkeeping is stripped unless the session
    /// was opened with <see cref="DocxSessionSettings.PersistAnchorIds"/>.
    /// </summary>
    public byte[] Save() => Save(_settings.PersistAnchorIds);

    /// <summary>
    /// Serialize with an explicit choice about the projector's <c>PtOpenXml:Unid</c> bookkeeping,
    /// overriding <see cref="DocxSessionSettings.PersistAnchorIds"/> for this call.
    /// </summary>
    /// <param name="persistAnchorIds">
    /// <c>true</c> keeps the Unid attributes in the output so a re-render of these bytes resolves to
    /// the SAME anchors the live session holds. That is an internal round-trip contract, not a
    /// document feature: the attributes are ~50 bytes each on every projected element (roughly 6x
    /// the file size of a real document), and while Word and LibreOffice both ignore them, bytes
    /// produced this way should not be handed to a user or written to disk as "the document".
    /// <c>false</c> — what a save-to-disk wants — strips them.
    /// </param>
    /// <remarks>
    /// The distinction exists because the two consumers genuinely differ: the browser editor's
    /// remount re-renders saved bytes and needs id stability across that hop, while its
    /// <c>save()</c> produces the file the user downloads. Making it a per-CALL choice rather than a
    /// session-wide setting is what keeps one from contaminating the other.
    /// </remarks>
    public byte[] Save(bool persistAnchorIds)
    {
        ThrowIfDisposed();

        if (persistAnchorIds)
        {
            // Flush every projected part's cached XDocument to its stream first.
            // Ops mutate the cached XDocument only; historically the per-op
            // projection rebuild flushed for them (scope.Part.PutXDocument in
            // BuildAnchorIndex), but that flush is now conditional on Unid
            // assignment — this path must not depend on it, or an op that changes
            // content without creating a new Unid (e.g. SetPageNumbering) could
            // serialize stale bytes.
            foreach (var part in EnumerateProjectedParts())
            {
                if (part.GetXDocument().Root is not null) part.PutXDocument();
            }
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return ZipUnixPermissionFixer.Fix(_stream.ToArray());
        }

        // Strip the internal PtOpenXml:Unid attributes before serializing — they're
        // projector bookkeeping, not OOXML schema, and on a real document the bloat
        // is significant (each Unid is ~50 bytes and the projector assigns one to
        // every descendant of every projected scope). We snapshot first so the
        // session's in-memory state can keep using Unids after the save completes;
        // Project() / Resolve() rely on them.
        var snapshot = TakeSnapshot();
        try
        {
            foreach (var part in EnumerateProjectedParts())
            {
                var xdoc = part.GetXDocument();
                if (xdoc.Root is null) continue;
                bool any = false;
                foreach (var el in xdoc.Root.DescendantsAndSelf())
                {
                    var attr = el.Attribute(PtOpenXml.Unid);
                    if (attr is not null) { attr.Remove(); any = true; }
                }
                if (any) part.PutXDocument();
            }
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return ZipUnixPermissionFixer.Fix(_stream.ToArray());
        }
        finally
        {
            RestoreSnapshot(snapshot);
        }
    }

    /// <summary>
    /// Enumerates every OOXML part the projector walks. Kept centralized so
    /// <see cref="Save"/> (Unid stripping) and any future part-level pass don't drift.
    /// </summary>
    /// <remarks>
    /// Includes every <see cref="CustomXmlPart"/> on the main document because
    /// callers like <see cref="ResolvePart"/> need to be able to look up any
    /// CustomXmlPart by URI. The snapshot/restore path uses
    /// <see cref="EnumerateProjectedPartsForSnapshot"/> instead, which narrows
    /// CustomXmlParts to the annotations part only — see that method for why.
    /// </remarks>
    private IEnumerable<OpenXmlPart> EnumerateProjectedParts()
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null) yield return main.WordprocessingCommentsPart;
        // Custom XML parts hold annotation metadata; include them so callers that
        // need to look up parts by URI (e.g. ResolvePart) can find them.
        foreach (var cx in main.CustomXmlParts) yield return cx;
    }

    /// <summary>
    /// Snapshot-scoped projected-part enumeration. Same as
    /// <see cref="EnumerateProjectedParts"/> for the structural parts, but narrows
    /// <see cref="OpenXmlPackaging.CustomXmlPart"/> enumeration to the Docxodus
    /// <em>annotations</em> CustomXmlPart only (identified by its root namespace
    /// via <see cref="Internal.AnnotationsCustomXml.Find"/>).
    /// </summary>
    /// <remarks>
    /// Why narrow here: <see cref="RestoreSnapshot"/> handles undo-time create/delete
    /// of CustomXmlParts via <c>AddCustomXmlPart(CustomXmlPartType.CustomXml)</c>,
    /// which hard-codes the content type and creates no
    /// <c>CustomXmlPropertiesPart</c> partner. That is correct for the annotations
    /// part but would silently corrupt other CustomXmlParts that Word/SharePoint
    /// rely on (SharePoint metadata, content-type-bound SDT data-binding parts,
    /// inkml, etc.) by re-creating them with the wrong content type and missing
    /// properties partner. Today no session op deletes non-annotation CustomXmlParts
    /// — narrowing here pre-empts the footgun before such an op is added.
    /// </remarks>
    private IEnumerable<OpenXmlPart> EnumerateProjectedPartsForSnapshot()
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null) yield return main.WordprocessingCommentsPart;
        // Comment-threading metadata parts: content is snapshot-scoped so reply/resolve writes
        // and DeleteBlock/RemoveComment pruning are undoable; create/delete reconciliation is
        // driven by DocumentSnapshot.CommentThreadingParts below.
        if (main.WordprocessingCommentsExPart is not null) yield return main.WordprocessingCommentsExPart;
        if (main.WordprocessingCommentsIdsPart is not null) yield return main.WordprocessingCommentsIdsPart;
        var annotationsPart = Internal.AnnotationsCustomXml.Find(_doc);
        if (annotationsPart is not null) yield return annotationsPart;
    }

    // ─── Tier A: text CRUD ────────────────────────────────────────────────

    public EditResult ReplaceText(string anchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceText requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        // Strip a leading auto-number prefix from the payload before parsing. The
        // projector emits "## Fourth The total number…" — auto-number from numPr
        // plus a space separator plus the run text — so an agent that echoes the
        // visible heading back as its replacement payload otherwise gets the
        // prefix applied twice (Word renders the auto-number AND the run text now
        // begins with "Fourth"). See DS091.
        markdownPayload = StripResolvedAutoNumberPrefix(element, markdownPayload);
        markdownPayload = MaybeApplySmartQuotes(markdownPayload);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                ApplyReplaceTextTracked(element, parsed.Blocks);
            }
            else
            {
                ApplyReplaceTextAccept(element, parsed.Blocks);
            }
            PromoteHyperlinkRelationships(element);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult DeleteBlock(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li" or "tbl" or "fn" or "en" or "cmt"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteBlock requires a block-level/footnote/endnote/comment anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        // Word reserves the TYPED footnote/endnote definitions (separator, continuationSeparator,
        // continuationNotice) for page-rendering scaffolding; they carry no user content and
        // removing one corrupts the document. Same predicate the projector filters on, so the two
        // can't drift over which types count as reserved.
        if (target.Anchor.Kind is "fn" or "en" && WmlToMarkdownConverter.IsBoilerplateNote(element))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"cannot delete a Word-reserved {target.Anchor.Kind} of type='{(string?)element.Attribute(W.type)}'",
                anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Tracked-change mode wraps removed runs in w:del — only meaningful for
            // body-level paragraph kinds. fn/en/cmt are structural definitions in
            // their own parts; "tracking" a definition deletion has no Word semantics,
            // so for those we always perform the structural delete.
            if (_trackedChanges == TrackedChangeMode.RenderInline
                && target.Anchor.Kind is "p" or "h" or "li")
            {
                WrapRunsInDel(element);
                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = new[] { target.Anchor },
                    Patch = PatchFor(target),
                };
            }

            // For fn/en/cmt: also remove every cross-reference (footnoteReference,
            // endnoteReference, commentReference/RangeStart/RangeEnd) anywhere in
            // the package that points at this definition's id. Otherwise Word
            // renders broken superscript references for the orphaned ids.
            if (target.Anchor.Kind is "fn" or "en" or "cmt")
            {
                var elementId = (string?)element.Attribute(W.id);
                if (!string.IsNullOrEmpty(elementId))
                    RemoveCrossReferences(target.Anchor.Kind, elementId);

                // For comments, also prune Word's threading metadata (commentsExtended /
                // commentsIds entries keyed by the definition paragraphs' w14:paraId) so a
                // removed comment leaves no dangling reply/resolve state. Lives here — not in
                // RemoveComment — so the generic DeleteBlock path gets it too.
                if (target.Anchor.Kind == "cmt")
                {
                    var paraIds = element.Elements(W.p)
                        .Select(p => (string?)p.Attribute(W14.paraId))
                        .Where(pid => !string.IsNullOrEmpty(pid))
                        .Select(pid => pid!)
                        .ToList();
                    Internal.CommentOps.PruneThreadingMetadata(_doc!, paraIds);
                }
            }

            // Collect descendant anchors before removal so the caller knows what's gone.
            var index = AnchorIndex();
            var removed = new List<Anchor> { target.Anchor };
            foreach (var d in element.Descendants())
            {
                var unid = (string?)d.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                foreach (var kv in index)
                {
                    if (kv.Value.Unid == unid && kv.Value.Unid != target.Unid)
                        removed.Add(kv.Value.Anchor);
                }
            }
            element.Remove();
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Removed = removed,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Deletes every top-level block-level element between <paramref name="fromAnchorId"/>
    /// (inclusive) and <paramref name="toAnchorIdExclusive"/> (exclusive) in document order.
    /// Both anchors must be block-level kinds (<c>p</c>, <c>h</c>, <c>li</c>, <c>tbl</c>),
    /// live in the same package part, and share a direct parent (no spanning into table
    /// cells or other nested containers). Records a single undo snapshot so
    /// <see cref="Undo"/> restores the entire range together.
    /// </summary>
    /// <remarks>
    /// In <see cref="TrackedChangeMode.RenderInline"/>, each paragraph in the range has
    /// its runs wrapped in <c>w:del</c> and its paragraph-mark marked deleted via
    /// <c>w:pPr/w:rPr/w:del</c>; each table row gets a <c>w:trPr/w:del</c> marker with
    /// its cell paragraphs wrapped recursively. Anchors stay live (<see cref="EditResult.Modified"/>
    /// instead of <see cref="EditResult.Removed"/>) so callers can re-address the same
    /// blocks before changes are accepted. Block-level elements other than <c>w:p</c>
    /// and <c>w:tbl</c> (e.g. <c>w:sdt</c>) are still structurally removed in this mode
    /// — issue #177 follow-up if a consumer needs them tracked.
    /// </remarks>
    public EditResult DeleteRange(string fromAnchorId, string toAnchorIdExclusive)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var fromTarget = FindAnchor(fromAnchorId);
        if (fromTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"from anchor not found: {fromAnchorId}", fromAnchorId);
        var toTarget = FindAnchor(toAnchorIdExclusive);
        if (toTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"to anchor not found: {toAnchorIdExclusive}", toAnchorIdExclusive);

        // Scope (package-part) check first — different parts can't form a contiguous
        // sibling range under any circumstance, even if the kinds look block-level.
        if (fromTarget.Anchor.Scope != toTarget.Anchor.Scope)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                $"DeleteRange anchors must live in the same package part; from={fromTarget.Anchor.Scope} to={toTarget.Anchor.Scope}",
                fromAnchorId);

        if (fromTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteRange requires block-level anchors; from kind={fromTarget.Anchor.Kind}", fromAnchorId);
        if (toTarget.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteRange requires block-level anchors; to kind={toTarget.Anchor.Kind}", toAnchorIdExclusive);

        var fromElement = fromTarget.Resolve(_doc!);
        var toElement = toTarget.Resolve(_doc!);
        if (fromElement is null || toElement is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", fromAnchorId);
        if (fromElement.Parent != toElement.Parent)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "DeleteRange anchors must share a direct parent (no spanning into nested containers)",
                fromAnchorId);

        return DeleteSiblingRangeCore(fromTarget, fromElement, toElement);
    }

    /// <summary>
    /// Deletes a heading and every block-level sibling under it, up to (but not including)
    /// the next heading at the same or higher level. If no such next heading exists, the
    /// section extends to the end of the parent (the heading and everything after it).
    /// </summary>
    /// <param name="headingAnchorId">Anchor id of the heading paragraph (kind must be <c>h</c>).</param>
    /// <remarks>
    /// "Level" is the same notion <see cref="WmlToMarkdownConverter"/> uses for the projection:
    /// <c>Heading1</c> = 1, <c>Heading2</c> = 2, etc.; <c>Title</c> = 1, <c>Subtitle</c> = 2.
    /// Tracked-change mode inherits <see cref="DeleteRange"/>'s behavior via the shared
    /// <c>DeleteSiblingRangeCore</c> helper: paragraphs and tables are wrapped in
    /// <c>w:del</c> markup rather than removed.
    /// </remarks>
    public EditResult DeleteSection(string headingAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var headingTarget = FindAnchor(headingAnchorId);
        if (headingTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"heading anchor not found: {headingAnchorId}", headingAnchorId);
        if (headingTarget.Anchor.Kind != "h")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteSection requires a heading anchor (kind=h); got kind={headingTarget.Anchor.Kind}",
                headingAnchorId);

        var headingElement = headingTarget.Resolve(_doc!);
        if (headingElement is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "heading element resolved null", headingAnchorId);

        int level = WmlToMarkdownConverter.HeadingLevel(headingElement);

        // Scan forward siblings for the next heading at level <= ours. If none, toElement
        // stays null and DeleteSiblingRangeCore will delete to the end of the parent.
        XElement? toElement = null;
        foreach (var sibling in headingElement.ElementsAfterSelf())
        {
            if (sibling.Name == W.p && WmlToMarkdownConverter.IsHeading(sibling)
                && WmlToMarkdownConverter.HeadingLevel(sibling) <= level)
            {
                toElement = sibling;
                break;
            }
        }

        return DeleteSiblingRangeCore(headingTarget, headingElement, toElement);
    }

    /// <summary>
    /// Shared core for <see cref="DeleteRange"/> and <see cref="DeleteSection"/>.
    /// Takes resolved XElement endpoints — <paramref name="toElementExclusive"/> may be
    /// <c>null</c> to mean "delete to the end of the parent". Records one snapshot and
    /// returns a single <see cref="EditResult"/> aggregating every removed anchor.
    /// </summary>
    private EditResult DeleteSiblingRangeCore(
        AnchorTarget anchorForPatchScope,
        XElement fromElement,
        XElement? toElementExclusive)
    {
        // Walk siblings from `fromElement` forward, accumulating elements to remove.
        var toRemove = new List<XElement>();
        var current = (XElement?)fromElement;
        while (current is not null && current != toElementExclusive)
        {
            toRemove.Add(current);
            current = current.ElementsAfterSelf().FirstOrDefault();
        }
        if (toElementExclusive is not null && current != toElementExclusive)
            return EditResult.Fail(EditErrorCode.InvalidPosition,
                "'to' anchor does not follow 'from' in document order",
                anchorForPatchScope.Anchor.Id);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var index = AnchorIndex();
            bool trackedChanges = _trackedChanges == TrackedChangeMode.RenderInline;

            if (trackedChanges)
            {
                // Tracked-change path: mark each block with w:del markup rather than
                // removing it. Anchors stay live in the document tree so callers can
                // re-address the same blocks before changes are accepted. Only the
                // top-level block anchors are reported as Modified — descendants stay
                // resolvable too, but enumerating them all would be noise (matches
                // DeleteBlock's single-anchor contract in tracked mode).
                var modified = new List<Anchor>();
                foreach (var el in toRemove)
                {
                    var elUnid = (string?)el.Attribute(PtOpenXml.Unid);
                    if (elUnid is not null)
                    {
                        foreach (var kv in index)
                            if (kv.Value.Unid == elUnid)
                                modified.Add(kv.Value.Anchor);
                    }
                    if (el.Name == W.p)
                        MarkParagraphAsTrackedDeleted(el);
                    else if (el.Name == W.tbl)
                        MarkTableAsTrackedDeleted(el);
                    else
                        // Block kinds beyond w:p/w:tbl (e.g. w:sdt) — v1 falls back
                        // to structural removal for these, per the issue-#177 docstring.
                        el.Remove();
                }
                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = modified,
                    Patch = PatchFor(anchorForPatchScope),
                };
            }

            var removed = new List<Anchor>();
            foreach (var el in toRemove)
            {
                // Collect this element's anchor plus every descendant anchor.
                CollectAnchorsForRemoval(el, index, removed);
                el.Remove();
            }
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Removed = removed,
                Patch = PatchFor(anchorForPatchScope),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorForPatchScope.Anchor.Id);
        }
    }

    private static void CollectAnchorsForRemoval(
        XElement el,
        IReadOnlyDictionary<string, AnchorTarget> index,
        List<Anchor> removed)
    {
        var elUnid = (string?)el.Attribute(PtOpenXml.Unid);
        if (elUnid is not null)
        {
            foreach (var kv in index)
                if (kv.Value.Unid == elUnid)
                    removed.Add(kv.Value.Anchor);
        }
        foreach (var desc in el.Descendants())
        {
            var dUnid = (string?)desc.Attribute(PtOpenXml.Unid);
            if (dUnid is null) continue;
            foreach (var kv in index)
                if (kv.Value.Unid == dUnid)
                    removed.Add(kv.Value.Anchor);
        }
    }

    /// <summary>
    /// Strips every cross-reference pointing at the named footnote/endnote/comment id
    /// from every part of the package that can hold one. For footnotes/endnotes that's
    /// just <c>w:footnoteReference</c>/<c>w:endnoteReference</c>; for comments it's the
    /// triple <c>w:commentReference</c> + <c>w:commentRangeStart</c> + <c>w:commentRangeEnd</c>
    /// — leaving any of the three behind makes Word render a broken comment marker.
    /// </summary>
    private void RemoveCrossReferences(string kind, string elementId)
    {
        XName referenceName = kind switch
        {
            "fn" => W.footnoteReference,
            "en" => W.endnoteReference,
            "cmt" => W.commentReference,
            _ => null!,
        };
        if (referenceName is null) return;

        foreach (var part in EnumerateProjectedParts())
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            bool any = false;
            foreach (var refEl in root.Descendants(referenceName)
                .Where(r => (string?)r.Attribute(W.id) == elementId).ToList())
            {
                var parentRun = refEl.Parent;
                refEl.Remove();
                any = true;
                // The reference was the only meaningful child of its <w:r> wrapper:
                // strip the run too so we don't leave behind an empty <w:r> with a
                // FootnoteReference run style (which Word renders as an empty styled
                // span — invisible but untidy and confusing to downstream tooling).
                RemoveEmptyRunIfNeeded(parentRun);
            }
            if (kind == "cmt")
            {
                foreach (var rangeEl in root.Descendants(W.commentRangeStart)
                    .Concat(root.Descendants(W.commentRangeEnd))
                    .Where(r => (string?)r.Attribute(W.id) == elementId).ToList())
                {
                    rangeEl.Remove();
                    any = true;
                }
            }
            if (any) part.PutXDocument();
        }
    }

    /// <summary>
    /// If <paramref name="run"/> is a <c>&lt;w:r&gt;</c> whose only remaining children
    /// are properties (<c>w:rPr</c>) — no text, no breaks, no fields, no other content —
    /// remove the run. Avoids leaving orphaned styled-empty spans after the meaningful
    /// child (a footnote/endnote reference) was stripped.
    /// </summary>
    private static void RemoveEmptyRunIfNeeded(XElement? run)
    {
        if (run is null || run.Name != W.r) return;
        foreach (var child in run.Elements())
        {
            if (child.Name == W.rPr) continue;
            return; // has meaningful content — keep the run
        }
        run.Remove();
    }

    // ─── Tier B: structural ops ──────────────────────────────────────────

    public EditResult InsertParagraph(string anchorId, Position pos, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
        if (parsed.Blocks.Count == 0)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "empty payload", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var created = new List<Anchor>();
            var newElements = new List<XElement>();
            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                // List items: try to inherit numbering from a sibling list item so the
                // payload actually projects as a bullet/numbered item. If no sibling
                // has numbering, the paragraph stays bare and the projector classifies
                // it as a plain "p" — which is what we report below.
                if (block.Kind is Internal.ParserBlockKind.BulletItem
                                or Internal.ParserBlockKind.OrderedItem)
                    TryInheritNumPrFromSibling(p, element);
                UnidHelper.AssignToSelfAndDescendants(p);
                newElements.Add(p);
                var unid = (string)p.Attribute(PtOpenXml.Unid)!;
                var kind = ClassifyParagraphKind(p);
                created.Add(new Anchor($"{kind}:{target.Anchor.Scope}:{unid}", kind, target.Anchor.Scope, unid));
            }

            if (pos == Position.Before)
            {
                foreach (var n in newElements) element.AddBeforeSelf(n);
            }
            else
            {
                XElement after = element;
                foreach (var n in newElements) { after.AddAfterSelf(n); after = n; }
            }

            foreach (var n in newElements) PromoteHyperlinkRelationships(n);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult SplitParagraph(string anchorId, int characterOffset)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SplitParagraph requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var totalText = ParagraphText(element);
        if (characterOffset < 0 || characterOffset > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} out of [0, {totalText.Length}]", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            var second = new XElement(W.p);
            XElement? newPPr = null;
            if (pPr is not null)
            {
                newPPr = new XElement(pPr);
                second.Add(newPPr);
            }

            // Split any run that straddles the offset (descends into hyperlinks/sdts),
            // then split any container (hyperlink) that still straddles, then move all
            // inline children + markers at-or-past the offset to `second`.
            SplitRunsAtOffset(element, characterOffset);
            SplitInlineContainersAtOffset(element, characterOffset);
            MoveInlineChildrenAfter(element, characterOffset, second);

            if (newPPr is not null)
            {
                // pageBreakBefore is a once-only property: the original paragraph keeps it; the new
                // paragraph must not inherit a second page break (matches Word clearing it on Enter).
                newPPr.Elements(W.pageBreakBefore).Remove();

                // An empty bordered paragraph is a horizontal rule; splitting it (Enter) must not
                // propagate the rule's border onto the fresh paragraph below — otherwise every Enter
                // stacks another rule and borders the body text (S-1 smoke-test finding 1a). A bordered
                // paragraph that HAS text keeps its border on both halves (boxed-block behavior).
                if (totalText.Length == 0)
                    newPPr.Elements(W.pBdr).Remove();

                // An empty Enter-at-end split starts a fresh paragraph. For a non-list paragraph whose
                // style declares a distinct next-paragraph style (e.g. Title/Heading -> Normal), rebase
                // the new paragraph onto that next style instead of cloning the heading: a clean pStyle,
                // dropping the heading-only direct props and the inherited paragraph-mark rPr that would
                // otherwise bake the heading's bold into freshly-typed text. List items are exempt so the
                // editor's Enter-continuation keeps the list going.
                bool emptySplit = characterOffset >= totalText.Length;
                bool isListItem = newPPr.Element(W.numPr) is not null;
                if (emptySplit && !isListItem)
                {
                    var curStyle = (string?)newPPr.Element(W.pStyle)?.Attribute(W.val);
                    var nextStyle = ResolveNextParagraphStyle(curStyle);
                    if (nextStyle is not null && !string.Equals(nextStyle, curStyle, StringComparison.Ordinal))
                    {
                        var rebuilt = new XElement(W.pPr,
                            new XElement(W.pStyle, new XAttribute(W.val, nextStyle)));
                        newPPr.ReplaceWith(rebuilt);
                        newPPr = rebuilt;
                    }
                }

                // Re-mint Unids on the new paragraph's property subtree so cloned property elements
                // (jc, ind, numPr, ...) don't carry the original's Unid onto a second element.
                foreach (var el in newPPr.DescendantsAndSelf())
                    el.Attributes(PtOpenXml.Unid).Remove();
            }

            UnidHelper.AssignToSelfAndDescendants(second);
            element.AddAfterSelf(second);

            var secondUnid = (string)second.Attribute(PtOpenXml.Unid)!;
            InvalidateProjectionCache();

            // The new paragraph's kind can differ from the original (Heading -> Normal via the
            // next-paragraph style), so resolve its anchor from the fresh projection rather than
            // assuming the original kind.
            var secondAnchor =
                AnchorForUnid(secondUnid, target.PartUri)
                ?? new Anchor(
                    $"{target.Anchor.Kind}:{target.Anchor.Scope}:{secondUnid}",
                    target.Anchor.Kind, target.Anchor.Scope, secondUnid);

            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Created = new[] { secondAnchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// The linked next-paragraph style (<c>w:style/w:next/@w:val</c>) for the given paragraph style
    /// id, read from the styles part; null when the id is empty/unknown or declares no next style.
    /// Read via <c>GetXDocument</c> (the same view <see cref="Internal.StyleFactory"/> writes through)
    /// so styles synthesized earlier in the session are visible.
    /// </summary>
    private string? ResolveNextParagraphStyle(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        var part = _doc?.MainDocumentPart?.StyleDefinitionsPart;
        var root = part?.GetXDocument().Root;
        if (root is null) return null;
        var style = root.Elements(W.style)
            .FirstOrDefault(st => (string?)st.Attribute(W.styleId) == styleId);
        return (string?)style?.Element(W.next)?.Attribute(W.val);
    }

    public EditResult MergeParagraphs(string firstAnchorId, string secondAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var firstTarget = FindAnchor(firstAnchorId);
        if (firstTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "first anchor not found", firstAnchorId);
        var secondTarget = FindAnchor(secondAnchorId);
        if (secondTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "second anchor not found", secondAnchorId);

        var firstEl = firstTarget.Resolve(_doc!);
        var secondEl = secondTarget.Resolve(_doc!);
        if (firstEl is null || secondEl is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null");

        if (!ReferenceEquals(firstEl.NextNode, secondEl))
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "MergeParagraphs requires second anchor to be the immediate next sibling of first");

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Insert a single-space separator if both sides end/start with non-whitespace.
            // Sentences from two paragraphs should not jam into one another.
            var firstTail = ParagraphText(firstEl);
            var secondHead = ParagraphText(secondEl);
            if (firstTail.Length > 0 && secondHead.Length > 0
                && !char.IsWhiteSpace(firstTail[^1])
                && !char.IsWhiteSpace(secondHead[0]))
            {
                firstEl.Add(new XElement(W.r,
                    new XElement(W.t,
                        new XAttribute(XNamespace.Xml + "space", "preserve"), " ")));
            }

            // Move every paragraph-level child from secondEl into firstEl in document
            // order — runs, hyperlinks, sdts, fldSimples, bookmarkStart/End, comment
            // range markers, etc. The old implementation only moved direct <w:r>
            // children which silently discarded everything else.
            foreach (var child in secondEl.Elements().ToList())
            {
                if (child.Name == W.pPr) continue; // second's pPr is dropped; first's wins
                child.Remove();
                firstEl.Add(child);
            }
            secondEl.Remove();
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { firstTarget.Anchor },
                Removed = new[] { secondTarget.Anchor },
                Patch = PatchFor(firstTarget),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    // ─── Raw escape hatch ────────────────────────────────────────────────

    public RawDocxOps Raw => _raw ??= new RawDocxOps(this);

    private static readonly HashSet<string> AllowedXmlNamespaces = new()
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",        // w:
        "http://schemas.openxmlformats.org/officeDocument/2006/math",          // m:
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", // wp:
        "http://schemas.openxmlformats.org/drawingml/2006/main",               // a:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", // r:
        "http://powertools.codeplex.com/2011",                                 // PtOpenXml (Unid)
    };

    internal string RawGetXmlInternal(string anchorId)
    {
        ThrowIfDisposed();
        var target = FindAnchor(anchorId);
        if (target is null)
            throw new ArgumentException($"anchor not found: {anchorId}");
        var element = target.Resolve(_doc!);
        return element?.ToString() ?? "";
    }

    /// <summary>
    /// The live, in-memory document backing this session. Exposed for read-only,
    /// in-assembly consumers (e.g. session-attached single-block HTML rendering) that
    /// must read the current tree/parts without the round-trip cost of <see cref="Save"/>.
    /// Do not mutate it outside the session's own edit methods.
    /// </summary>
    internal WordprocessingDocument LiveDocument
    {
        get
        {
            ThrowIfDisposed();
            return _doc!;
        }
    }

    // Cached formatting "shell" for session-attached single-block rendering (see
    // Internal.HtmlConversionOps.RenderBlockHtml). A serialized throwaway .docx holding the
    // formatting parts (styles/numbering/theme/fontTable/settings) + an empty body, built ONCE and
    // reused across renders so a keystroke commit doesn't re-clone the (potentially huge) style
    // gallery every time. HtmlConversionOps owns these; it rebuilds the shell when
    // <see cref="RenderShellSignature"/> (a cheap content signature of the formatting parts) changes
    // — i.e. when a format op adds a style / numbering level. Text edits never touch those parts, so
    // the shell survives normal typing. Disposed implicitly with the session (plain GC).
    internal byte[]? RenderShellBytes;
    internal long RenderShellSignature;

    internal EditResult RawInsertXmlInternal(string anchorId, Position pos, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var (parsedXml, err) = ParseRawXml(xml);
        if (parsedXml is null)
            return new EditResult { Success = false, Error = err! with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        int baselineErrors = _settings.ValidateRawOps ? CountRealValidationErrors() : 0;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            UnidHelper.AssignToSelfAndDescendants(parsedXml);
            if (pos == Position.Before) element.AddBeforeSelf(parsedXml);
            else element.AddAfterSelf(parsedXml);

            if (_settings.ValidateRawOps && CountRealValidationErrors() > baselineErrors)
            {
                var preOp = _history.PopForUndo();
                if (preOp.ok) RestoreSnapshot(preOp.snapshot);
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            InvalidateProjectionCache();
            var freshIndex = AnchorIndex();
            var created = new List<Anchor>();
            foreach (var unid in CollectUnids(parsedXml))
            {
                var hit = AnchorForUnid(unid, target.PartUri);
                if (hit is { } h) created.Add(h);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    internal EditResult RawReplaceXmlInternal(string anchorId, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var (parsedXml, err) = ParseRawXml(xml);
        if (parsedXml is null)
            return new EditResult { Success = false, Error = err! with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        int baselineErrors = _settings.ValidateRawOps ? CountRealValidationErrors() : 0;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            UnidHelper.AssignToSelfAndDescendants(parsedXml);
            element.ReplaceWith(parsedXml);

            if (_settings.ValidateRawOps && CountRealValidationErrors() > baselineErrors)
            {
                var preOp = _history.PopForUndo();
                if (preOp.ok) RestoreSnapshot(preOp.snapshot);
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            InvalidateProjectionCache();
            var freshIndex = AnchorIndex();
            var newUnids = CollectUnids(parsedXml).ToHashSet();

            // Classify by Unid set membership: the documented Get→mutate→Replace
            // recipe preserves Unids, so the target anchor must surface as
            // Modified (not as a phantom Removed-then-Created pair). When the
            // replacement XML has fresh Unids — because the caller authored it
            // from scratch — the target is genuinely Removed and the new
            // element(s) are Created. See DS092 / DS092b.
            var modified = new List<Anchor>();
            var removed = new List<Anchor>();
            var created = new List<Anchor>();

            if (newUnids.Contains(target.Unid))
            {
                var hit = AnchorForUnid(target.Unid, target.PartUri);
                if (hit is { } h) modified.Add(h);
            }
            else
            {
                removed.Add(target.Anchor);
            }
            foreach (var unid in newUnids)
            {
                if (unid == target.Unid) continue;
                var hit = AnchorForUnid(unid, target.PartUri);
                if (hit is { } h) created.Add(h);
            }

            return new EditResult
            {
                Success = true,
                Removed = removed,
                Created = created,
                Modified = modified,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private static (XElement? parsed, EditError? err) ParseRawXml(string xml)
    {
        try
        {
            var x = XElement.Parse(xml);
            foreach (var el in x.DescendantsAndSelf())
            {
                var ns = el.Name.NamespaceName;
                if (!string.IsNullOrEmpty(ns) && !AllowedXmlNamespaces.Contains(ns))
                    return (null, new EditError(EditErrorCode.DisallowedNamespace,
                        $"disallowed namespace: {ns}"));
            }
            return (x, null);
        }
        catch (System.Xml.XmlException ex)
        {
            return (null, new EditError(EditErrorCode.MalformedXml, ex.Message));
        }
    }

    private static IEnumerable<string> CollectUnids(XElement root)
    {
        foreach (var el in root.DescendantsAndSelf())
        {
            var unid = (string?)el.Attribute(PtOpenXml.Unid);
            if (unid is not null) yield return unid;
        }
    }

    // PtOpenXml:Unid is an internal-only attribute added by the projector for anchor
    // addressing; it is not in the OOXML schema, so the validator will emit
    // Sch_UndeclaredAttribute for every occurrence. Filter those out before counting.
    //
    // Mutations operate directly on the part's in-memory XDocument; the validator
    // reads the typed OOXML object model, which is hydrated from the part stream.
    // Flush the XDocument back to the stream first so the validator sees the
    // current state instead of the original document.
    private int CountRealValidationErrors()
    {
        _doc!.MainDocumentPart!.PutXDocument();
        var v = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        return v.Validate(_doc!)
            .Count(e => !(e.Description ?? string.Empty)
                .Contains("http://powertools.codeplex.com/2011", StringComparison.Ordinal));
    }

    // ─── Tier D: table cell content ──────────────────────────────────────

    public EditResult ReplaceCellContent(string cellAnchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(cellAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", cellAnchorId);
        if (target.Anchor.Kind != "tc")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ReplaceCellContent requires a cell anchor", cellAnchorId);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, cellAnchorId);

        var cell = target.Resolve(_doc!);
        if (cell is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", cellAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var p in cell.Elements(W.p).ToList()) p.Remove();

            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                UnidHelper.AssignToSelfAndDescendants(p);
                cell.Add(p);
                PromoteHyperlinkRelationships(p);
            }
            // A table cell must contain at least one paragraph per OOXML schema.
            if (!cell.Elements(W.p).Any())
                cell.Add(new XElement(W.p));

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    // ─── Tier C: formatting ──────────────────────────────────────────────

    /// <summary>
    /// Convenience: find <paramref name="substring"/> in the anchor's flat text and apply
    /// <paramref name="op"/> to the first occurrence. Eliminates the offset-arithmetic
    /// trap where an auto-number prefix shifts the visible text vs the run-text indices
    /// the underlying <see cref="ApplyFormat(string, CharSpan?, FormatOp)"/> overload
    /// expects — see issue #138. Named distinctly (rather than overloading) so existing
    /// <c>ApplyFormat(anchor, null, op)</c> calls (whole-paragraph format) stay
    /// unambiguous to the C# overload resolver.
    /// </summary>
    public EditResult ApplyFormatToSubstring(string anchorId, string substring, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (string.IsNullOrEmpty(substring))
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "substring must be non-empty", anchorId);

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ApplyFormat requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var map = Internal.RunTextMap.Build(element);
        var idx = map.FlatText.IndexOf(substring, StringComparison.Ordinal);
        if (idx < 0) return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
            $"substring not found in anchor's text", anchorId);

        return ApplyFormat(anchorId, new CharSpan(idx, substring.Length), op);
    }

    /// <summary>
    /// Convenience: apply <paramref name="op"/> to the exact span covered by a
    /// <see cref="TextMatch"/> (typically from <see cref="Grep"/>). The match's
    /// <see cref="TextMatch.EnclosingAnchor"/> + <see cref="TextMatch.Span"/> address
    /// one specific occurrence even when several identical needles share the same block.
    /// </summary>
    public EditResult ApplyFormat(TextMatch match, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");
        return ApplyFormat(
            match.EnclosingAnchor.Anchor.Id,
            new CharSpan(match.Span.Start, match.Span.Length),
            op);
    }

    public EditResult ApplyFormat(string anchorId, CharSpan? span, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (op is null) return EditResult.Fail(EditErrorCode.MalformedMarkdown, "null format op", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ApplyFormat requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var totalText = ParagraphText(element);
        var actualSpan = span ?? new CharSpan(0, totalText.Length);
        if (actualSpan.Start < 0 || actualSpan.Length < 0 ||
            actualSpan.Start + actualSpan.Length > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span [{actualSpan.Start},{actualSpan.Start + actualSpan.Length}) out of [0,{totalText.Length})", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Inline code references a "Code" character style by id; ensure it actually
            // exists so the run renders monospace instead of pointing at a phantom style.
            if (op.Code is true) Internal.StyleFactory.EnsureCodeCharacterStyle(_doc);

            SplitRunsAtOffset(element, actualSpan.Start);
            SplitRunsAtOffset(element, actualSpan.Start + actualSpan.Length);

            var trackFormatChanges = _trackedChanges == TrackedChangeMode.RenderInline;
            var revisionAuthor = _revisionAuthor ?? "docxodus";
            // Every run touched by one ApplyFormat call belongs to the same user action.
            // Give its per-run rPrChange markers one timestamp for coherent attribution
            // and grouping. The value is monotonic at tick precision so two adjacent,
            // same-author ApplyFormat calls remain independently selectable even when
            // the wall clock would otherwise stamp them in the same second.
            var revisionDate = trackFormatChanges
                ? NextTrackedFormatRevisionDate()
                : null;

            int consumed = 0;
            foreach (var run in InlineRuns(element).ToList())
            {
                var runText = RunText(run);
                int runStart = consumed;
                int runEnd = consumed + runText.Length;
                consumed = runEnd;
                if (runEnd <= actualSpan.Start || runStart >= actualSpan.Start + actualSpan.Length) continue;
                if (trackFormatChanges)
                    ApplyFormatToRunTracked(run, op, revisionAuthor, revisionDate!);
                else
                    ApplyFormatToRun(run, op);
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private string NextTrackedFormatRevisionDate()
    {
        while (true)
        {
            var observed = System.Threading.Interlocked.Read(ref _lastFormatRevisionTicks);
            var now = DateTime.UtcNow.Ticks;
            var next = now > observed ? now : observed + 1;
            if (System.Threading.Interlocked.CompareExchange(
                    ref _lastFormatRevisionTicks, next, observed) == observed)
            {
                return new DateTime(next, DateTimeKind.Utc).ToString(
                    "yyyy-MM-ddTHH:mm:ss.fffffffZ",
                    System.Globalization.CultureInfo.InvariantCulture);
            }
        }
    }

    public EditResult SetParagraphStyle(string anchorId, string styleId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetParagraphStyle requires a paragraph anchor", anchorId);

        // Find-or-create well-known built-in styles (Title, Subtitle, Heading1-9) the document
        // hasn't defined yet, so applying one works instead of silently failing. Mirrors the inline
        // "Code" character style. A truly unknown custom id is left untouched and still rejected.
        if (!Internal.StyleFactory.EnsureParagraphStyle(_doc!, styleId))
            return EditResult.Fail(EditErrorCode.UnknownStyle, $"style id not found: {styleId}", anchorId);

        var element = target.Resolve(_doc);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            pPr.Element(W.pStyle)?.Remove();
            pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));

            InvalidateProjectionCache();
            // Anchor kind may have flipped (e.g., p → h); look it up in the fresh index.
            var freshIndex = AnchorIndex();
            var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    // CT_PPr child schema order (subset covering what we insert). w:pPr children must
    // appear in this sequence or Word treats the file as needing repair.
    private static readonly string[] PPrChildOrder =
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
        "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
        "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
        "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
        "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
        "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
    };

    /// <summary>Return the existing w:pPr child of <paramref name="name"/> (attributes intact),
    /// or slot a new empty one in at its correct CT_PPr position.</summary>
    private static XElement GetOrCreatePPrChild(XElement pPr, XName name)
    {
        var child = pPr.Element(name);
        if (child is null)
        {
            child = new XElement(name);
            SetPPrChildInOrder(pPr, child);
        }
        return child;
    }

    /// <summary>Insert (replacing any existing) a w:pPr child at its correct CT_PPr position.</summary>
    private static void SetPPrChildInOrder(XElement pPr, XElement child)
    {
        pPr.Elements(child.Name).Remove();
        int idx = Array.IndexOf(PPrChildOrder, child.Name.LocalName);
        XElement? after = null;
        foreach (var e in pPr.Elements())
        {
            int ei = Array.IndexOf(PPrChildOrder, e.Name.LocalName);
            if (ei >= 0 && ei < idx) after = e;
            else if (ei >= idx) break;
        }
        if (after is null) pPr.AddFirst(child);
        else after.AddAfterSelf(child);
    }

    // CT_PBdr child schema order. w:pBdr edges must appear in this sequence.
    private static readonly string[] PBdrEdgeOrder = { "top", "left", "bottom", "right", "between", "bar" };

    private static XElement BorderEdgeElement(XName edgeName, ParagraphBorderEdge edge) =>
        new XElement(edgeName,
            new XAttribute(W.val, string.IsNullOrEmpty(edge.Style) ? "single" : edge.Style),
            new XAttribute(W.sz, edge.Size ?? 6),
            new XAttribute(W.space, edge.Space ?? 1),
            new XAttribute(W.color, string.IsNullOrEmpty(edge.Color) ? "auto" : edge.Color));

    /// <summary>Insert/replace a single <c>w:pBdr</c> edge, keeping CT_PBdr child order.</summary>
    private static void SetBorderEdgeInOrder(XElement pBdr, XName edgeName, XElement edge)
    {
        pBdr.Elements(edgeName).Remove();
        int idx = Array.IndexOf(PBdrEdgeOrder, edgeName.LocalName);
        XElement? after = null;
        foreach (var e in pBdr.Elements())
        {
            int ei = Array.IndexOf(PBdrEdgeOrder, e.Name.LocalName);
            if (ei >= 0 && ei < idx) after = e;
            else if (ei >= idx) break;
        }
        if (after is null) pBdr.AddFirst(edge);
        else after.AddAfterSelf(edge);
    }

    /// <summary>Apply top/bottom border edges (and an optional clear) to a paragraph's pPr, in place.</summary>
    private static void ApplyParagraphBorders(XElement pPr, ParagraphBorderEdge? top, ParagraphBorderEdge? bottom, bool clear)
    {
        if (clear) pPr.Element(W.pBdr)?.Remove();
        if (top is null && bottom is null) return;
        var pBdr = pPr.Element(W.pBdr);
        bool isNew = pBdr is null;
        pBdr ??= new XElement(W.pBdr);
        if (top is not null) SetBorderEdgeInOrder(pBdr, W.top, BorderEdgeElement(W.top, top));
        if (bottom is not null) SetBorderEdgeInOrder(pBdr, W.bottom, BorderEdgeElement(W.bottom, bottom));
        if (isNew) SetPPrChildInOrder(pPr, pBdr);
    }

    /// <summary>
    /// Set paragraph-level formatting (alignment, indent delta, first-line/hanging indent,
    /// before/after/line spacing, page-break-before, borders) on the paragraph the anchor
    /// names. Only the non-null fields of <paramref name="op"/> change.
    /// </summary>
    public EditResult SetParagraphFormat(string anchorId, ParagraphFormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetParagraphFormat requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        if (op.FirstLineIndent is not null && op.HangingIndent is not null)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "firstLineIndent and hangingIndent are mutually exclusive (w:ind holds one or the other)", anchorId);
        if (op.FirstLineIndent is < 0 || op.HangingIndent is < 0 ||
            op.SpacingBefore is < 0 || op.SpacingAfter is < 0 || op.LineSpacing is < 0)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "indent/spacing values are unsigned twips and must be >= 0", anchorId);
        if (op.LineSpacingRule is not null && op.LineSpacing is null)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "lineSpacingRule requires lineSpacing (w:lineRule qualifies w:line)", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }

            if (op.Alignment is { } align)
            {
                var val = align switch
                {
                    ParagraphAlignment.Left => "left",
                    ParagraphAlignment.Center => "center",
                    ParagraphAlignment.Right => "right",
                    ParagraphAlignment.Justify => "both",
                    _ => "left",
                };
                SetPPrChildInOrder(pPr, new XElement(W.jc, new XAttribute(W.val, val)));
            }

            if (op.PageBreakBefore is { } pbb)
            {
                pPr.Element(W.pageBreakBefore)?.Remove();
                if (pbb) SetPPrChildInOrder(pPr, new XElement(W.pageBreakBefore));
            }

            if (op.IndentDelta is { } delta && delta != 0)
            {
                var ind = pPr.Element(W.ind);
                // Parse the current left indent tolerantly: documents exported by Google Docs (and
                // others) emit non-integer twips like w:left="12.996749877929688", which a bare
                // (int?) cast rejects with a FormatException. AttributeToTwips is the same helper the
                // HTML converter uses (decimal → truncate), so we read what the doc renders and write
                // back a clean integer.
                int cur = ind is null ? 0 : WordprocessingMLUtil.AttributeToTwips(ind.Attribute(W.left)) ?? 0;
                int next = Math.Max(0, cur + delta);
                if (ind is null)
                {
                    ind = new XElement(W.ind);
                    SetPPrChildInOrder(pPr, ind);
                }
                ind.SetAttributeValue(W.left, next);
            }

            // firstLine/hanging share one w:ind slot in Word: writing either evicts the other
            // (validation above already rejected an op carrying both).
            if (op.FirstLineIndent is { } firstLine)
            {
                var ind = GetOrCreatePPrChild(pPr, W.ind);
                ind.SetAttributeValue(W.firstLine, firstLine);
                ind.SetAttributeValue(W.hanging, null);
            }
            if (op.HangingIndent is { } hanging)
            {
                var ind = GetOrCreatePPrChild(pPr, W.ind);
                ind.SetAttributeValue(W.hanging, hanging);
                ind.SetAttributeValue(W.firstLine, null);
            }

            if (op.SpacingBefore is not null || op.SpacingAfter is not null || op.LineSpacing is not null)
            {
                var spacing = GetOrCreatePPrChild(pPr, W.spacing);
                // A direct beforeAutospacing/afterAutospacing flag makes Word ignore the explicit
                // value, so writing one clears the matching flag — Word's own Paragraph dialog
                // does the same when a typed value replaces "Auto".
                if (op.SpacingBefore is { } before)
                {
                    spacing.SetAttributeValue(W.before, before);
                    spacing.SetAttributeValue(W.beforeAutospacing, null);
                }
                if (op.SpacingAfter is { } after)
                {
                    spacing.SetAttributeValue(W.after, after);
                    spacing.SetAttributeValue(W.afterAutospacing, null);
                }
                if (op.LineSpacing is { } line)
                {
                    spacing.SetAttributeValue(W.line, line);
                    spacing.SetAttributeValue(W.lineRule, (op.LineSpacingRule ?? LineSpacingRule.Auto) switch
                    {
                        LineSpacingRule.Exact => "exact",
                        LineSpacingRule.AtLeast => "atLeast",
                        _ => "auto",
                    });
                }
            }

            if (op.ClearBorders is true || op.TopBorder is not null || op.BottomBorder is not null)
                ApplyParagraphBorders(pPr, op.TopBorder, op.BottomBorder, op.ClearBorders is true);

            InvalidateProjectionCache();
            var freshIndex = AnchorIndex();
            var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Insert an empty paragraph carrying a bottom border — an S-1-style horizontal rule —
    /// before/after the block named by <paramref name="anchorId"/>. <paramref name="rule"/>
    /// styles the line (default: a single 12-eighths ≈1.5pt black rule).
    /// </summary>
    /// <summary>
    /// Mint a complete, blank single-paragraph DOCX (Normal style, doc defaults, settings, and a
    /// US-Letter portrait section) as bytes — a "New document" seed for editors that draft from
    /// scratch. The result opens cleanly in Word and as a <see cref="DocxSession"/>.
    /// </summary>
    public static byte[] CreateBlankDocxBytes() => Internal.BlankDocumentFactory.CreateBytes();

    /// <summary>
    /// Insert an empty paragraph carrying a bottom border — an S-1-style horizontal rule —
    /// before/after the block named by <paramref name="anchorId"/>. <paramref name="rule"/>
    /// styles the line (default: a single 12-eighths ≈1.5pt black rule).
    /// </summary>
    public EditResult InsertHorizontalRule(string anchorId, Position pos, ParagraphBorderEdge? rule = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var edge = rule ?? new ParagraphBorderEdge { Style = "single", Size = 12, Color = "auto" };
            var pPr = new XElement(W.pPr);
            ApplyParagraphBorders(pPr, top: null, bottom: edge, clear: false);
            var p = new XElement(W.p, pPr);
            UnidHelper.AssignToSelfAndDescendants(p);

            if (pos == Position.Before) element.AddBeforeSelf(p);
            else element.AddAfterSelf(p);

            var unid = (string)p.Attribute(PtOpenXml.Unid)!;
            InvalidateProjectionCache();
            var created = AnchorForUnid(unid, target.PartUri)
                ?? new Anchor($"p:{target.Anchor.Scope}:{unid}", "p", target.Anchor.Scope, unid);

            return new EditResult
            {
                Success = true,
                Created = new[] { created },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    // ─── Headers / footers / page-number fields ───────────────────────────────
    //
    // Author the per-section running header/footer stories (which live in their own OOXML
    // parts, outside the body) and page-number fields. SetHeaderText/SetFooterText are
    // addressed by ANY body block in the target section — the governing w:sectPr is resolved
    // the same way GetSectionInfo resolves it (a mid-document section break, else the body's
    // trailing sectPr, creating one if the body has none). The created header/footer paragraph
    // anchors come back in EditResult.Created with a hdr{N}/ftr{N} scope, so a page-number field
    // can then be inserted into them with InsertPageNumberField. Undo/redo of the part creation
    // is handled by the header/footer reconcile in RestoreSnapshot.

    /// <summary>
    /// Set the running <b>header</b> story for the section that owns <paramref name="anchorId"/>
    /// (any body block in that section) to <paramref name="markdownPayload"/>. Creates the header
    /// part, its relationship, and the <c>w:headerReference</c> on the section if the story of the
    /// requested <paramref name="kind"/> does not exist yet; otherwise replaces that story's content.
    /// An empty payload yields a single empty header paragraph. <see cref="HeaderFooterKind.First"/>
    /// sets the section's <c>w:titlePg</c>; <see cref="HeaderFooterKind.Even"/> sets
    /// <c>w:evenAndOddHeaders</c> in the settings part. Returns the created header-paragraph anchors
    /// (scope <c>hdr{N}</c>) in <see cref="EditResult.Created"/>.
    /// </summary>
    public EditResult SetHeaderText(string anchorId, HeaderFooterKind kind, string markdownPayload)
        => SetHeaderFooterText(isHeader: true, anchorId, kind, markdownPayload);

    /// <summary>
    /// Set the running <b>footer</b> story for the section that owns <paramref name="anchorId"/>.
    /// Behaves exactly like <see cref="SetHeaderText"/> but for the footer part / <c>w:footerReference</c>;
    /// the created footer-paragraph anchors (scope <c>ftr{N}</c>) come back in
    /// <see cref="EditResult.Created"/> — insert a page number into one with
    /// <see cref="InsertPageNumberField"/>.
    /// </summary>
    public EditResult SetFooterText(string anchorId, HeaderFooterKind kind, string markdownPayload)
        => SetHeaderFooterText(isHeader: false, anchorId, kind, markdownPayload);

    /// <summary>
    /// Ensure Word will actually RENDER the <paramref name="kind"/> header/footer stories of the
    /// section that owns <paramref name="anchorId"/> (any body block, resolved as
    /// <see cref="GetSectionInfo"/> resolves it): sets <c>w:titlePg</c> for
    /// <see cref="HeaderFooterKind.First"/> and the document-global <c>w:evenAndOddHeaders</c>
    /// for <see cref="HeaderFooterKind.Even"/>. <see cref="HeaderFooterKind.Default"/> needs no
    /// flag and is a successful no-op. Idempotent.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="SetHeaderText"/>/<see cref="SetFooterText"/> set these flags as a side effect of
    /// writing content, which covers authoring a story from scratch. It does NOT cover a document
    /// that already carries a first/even reference with the flag absent — Word writes exactly that
    /// when "Different first page" / "Different odd &amp; even pages" is turned back off, leaving
    /// the part behind. Editing such a story through the anchor-addressed text ops then produces a
    /// document whose header content is present but invisible. An editor offering a
    /// first/even story selector needs this as its own operation, because the flags belong to the
    /// SECTION, not to a content write.
    /// </para>
    /// <para>
    /// See <see cref="SetHeaderText"/> for the <c>w:evenAndOddHeaders</c> caveat: it is
    /// document-global and governs footers too, so enabling it without an even FOOTER means even
    /// pages show no footer at all.
    /// </para>
    /// </remarks>
    public EditResult EnsureHeaderFooterVisible(string anchorId, HeaderFooterKind kind)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "EnsureHeaderFooterVisible requires a body block anchor (the section the header/footer belongs to)",
                anchorId);
        if (kind == HeaderFooterKind.Default)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
        if (sectPr is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no governing section properties", anchorId);

        // Already set? Return BEFORE snapshotting. A UI that calls this on every kind selection
        // (the editor's band does) would otherwise push a no-op snapshot per click into the
        // bounded undo ring and evict the user's real history.
        bool alreadySet = kind == HeaderFooterKind.First
            ? sectPr.Element(W.titlePg) is not null
            : main.DocumentSettingsPart?.GetXDocument().Root?.Element(W.evenAndOddHeaders) is not null;
        if (alreadySet)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (kind == HeaderFooterKind.First) InsertSectPrTitlePg(sectPr);
            else WordprocessingMLUtil.EnsureEvenAndOddHeaders(main);
            InvalidateProjectionCache();
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private EditResult SetHeaderFooterText(bool isHeader, string anchorId, HeaderFooterKind kind, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "SetHeaderText/SetFooterText require a body block anchor (the section the header/footer belongs to)", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var body = element.AncestorsAndSelf(W.body).FirstOrDefault();
        if (body is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "anchor is not in the document body", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        // Parse the payload into paragraphs (empty payload ⇒ one empty paragraph), then apply the
        // built-in Header/Footer style so the paragraphs inherit Word's centre/right tab stops.
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));
        ApplyHeaderFooterStyle(paras, isHeader);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
            if (sectPr is null)
            {
                // A body with no section properties at all — synthesize the document-final section.
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
            }

            var refName = isHeader ? W.headerReference : W.footerReference;
            var typeVal = HeaderFooterTypeValue(kind);

            // Reuse the same-kind reference's part if it resolves to the right part type; otherwise
            // add a fresh part and reference (dropping any stale/mismatched same-kind reference).
            var existingRef = sectPr.Elements(refName)
                .FirstOrDefault(r => (string?)r.Attribute(W.type) == typeVal);
            OpenXmlPart? reuse = null;
            if (existingRef is not null && (string?)existingRef.Attribute(R.id) is { } rid)
                foreach (var pp in main.Parts)
                    if (pp.RelationshipId == rid) { reuse = pp.OpenXmlPart; break; }
            bool typeMatches = isHeader ? reuse is HeaderPart : reuse is FooterPart;

            OpenXmlPart part;
            if (reuse is not null && typeMatches)
            {
                part = reuse;
            }
            else
            {
                part = isHeader ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
                existingRef?.Remove();
                // Header/footer references lead the CT_SectPr sequence, so AddFirst is schema-ordered.
                sectPr.AddFirst(new XElement(refName,
                    new XAttribute(W.type, typeVal),
                    new XAttribute(R.id, main.GetIdOfPart(part))));
            }

            // Stamp Unids so the new paragraphs can be reported as anchors after re-projection.
            foreach (var p in paras) UnidHelper.AssignToSelfAndDescendants(p);

            var newRoot = new XElement(isHeader ? W.hdr : W.ftr,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "r", R.r),
                paras);
            part.PutXDocument(new XDocument(newRoot));

            // Visibility flags so Word actually shows the First/Even stories.
            if (kind == HeaderFooterKind.First && sectPr.Element(W.titlePg) is null)
                InsertSectPrTitlePg(sectPr);
            else if (kind == HeaderFooterKind.Even)
                WordprocessingMLUtil.EnsureEvenAndOddHeaders(main);

            InvalidateProjectionCache();
            var index = AnchorIndex();
            var created = new List<Anchor>();
            foreach (var p in paras)
            {
                var unid = (string?)p.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                var t = AnchorForUnid(unid, PartUriOf(p));
                if (t is { } a) created.Add(a);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Append a page-number field to the paragraph named by <paramref name="anchorId"/> — typically a
    /// header/footer paragraph (e.g. one returned by <see cref="SetFooterText"/>), though any paragraph
    /// is accepted. <see cref="PageNumberField.CurrentPage"/> emits a <c>PAGE</c> field,
    /// <see cref="PageNumberField.TotalPages"/> a <c>NUMPAGES</c> field, both as a native Word complex
    /// field (<c>fldChar</c>/<c>instrText</c>) with a cached result. Center the number by setting the
    /// paragraph alignment (<see cref="SetParagraphFormat"/>) or by relying on the Header/Footer style's
    /// centre tab. Returns the affected paragraph anchor in <see cref="EditResult.Modified"/>.
    /// </summary>
    /// <param name="format">
    /// Optional per-field number format, written as the field's <c>\*</c> general-formatting switch
    /// (<c>PAGE \* roman</c> → <c>i, ii, iii</c>). <c>null</c> — the default — emits a plain field,
    /// which is what Word inserts and what follows the SECTION's format
    /// (<see cref="SetPageNumbering"/>). Prefer the section setting for ordinary page numbering; a
    /// switch here OVERRIDES it for this one field and keeps overriding it if the section later
    /// changes. <see cref="NumberFormat.Bullet"/> is rejected.
    /// </param>
    public EditResult InsertPageNumberField(
        string anchorId,
        PageNumberField field = PageNumberField.CurrentPage,
        NumberFormat? format = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (format is { } f && !Internal.NumberFormats.IsPageNumberFormat(f))
            return EditResult.Fail(EditErrorCode.InvalidPageNumbering,
                $"{f} cannot format a page number", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        if (element.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "InsertPageNumberField requires a paragraph anchor", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var r in BuildPageNumberFieldRuns(field, format))
            {
                UnidHelper.AssignToSelfAndDescendants(r);
                element.Add(r);
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>Map <see cref="HeaderFooterKind"/> to the OOXML <c>w:type</c> token.</summary>
    private static string HeaderFooterTypeValue(HeaderFooterKind kind) => kind switch
    {
        HeaderFooterKind.First => "first",
        HeaderFooterKind.Even => "even",
        _ => "default",
    };

    /// <summary>Give every header/footer paragraph that carries no explicit <c>w:pStyle</c> the
    /// built-in Header/Footer style, so it inherits Word's centre-of-page and right-margin tab stops.</summary>
    private static void ApplyHeaderFooterStyle(List<XElement> paras, bool isHeader)
    {
        var styleId = isHeader ? "Header" : "Footer";
        foreach (var p in paras)
        {
            var pPr = p.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); p.AddFirst(pPr); }
            if (pPr.Element(W.pStyle) is null)
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));
        }
    }

    /// <summary>The runs of a native complex page-number field (PAGE / NUMPAGES) with a cached
    /// result — the form Word emits, so it renders and updates like a hand-authored field. With a
    /// <paramref name="format"/> the instruction carries the <c>\*</c> general-formatting switch and
    /// the cached result is page 1 rendered in that format, so a renderer that does not recompute
    /// fields shows a number consistent with the switch instead of always "1".</summary>
    private static XElement[] BuildPageNumberFieldRuns(PageNumberField field, NumberFormat? format)
    {
        var name = field == PageNumberField.TotalPages ? "NUMPAGES" : "PAGE";
        var instr = format is { } f && Internal.NumberFormats.ToFieldSwitch(f) is { } sw
            ? $" {name} \\* {sw} "
            : $" {name} ";
        var cached = format is { } cf ? Internal.NumberFormats.Render(1, cf) : "1";
        return new[]
        {
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
            new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), instr)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
            new XElement(W.r, new XElement(W.t, cached)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))),
        };
    }

    // ─── Section page numbering (w:pgNumType, issue #277) ─────────────────────
    //
    // The section-level half of page numbering: which number the section starts at and which format
    // its pages are numbered in. A plain PAGE field renders through this, so it — not the field —
    // is the normal place to say "front matter is i, ii, iii and the body restarts at 1".
    // Addressed by any body block in the target section, resolving the governing w:sectPr exactly
    // as GetSectionInfo does.

    /// <summary>
    /// Set the page-numbering properties (<c>w:pgNumType</c>) of the section that owns
    /// <paramref name="anchorId"/> (any body block in that section). Null fields on
    /// <paramref name="op"/> leave that attribute alone, so a caller can set the start without
    /// disturbing the format and vice versa. Creates the element, and a trailing <c>w:sectPr</c>,
    /// if absent. Idempotent.
    /// </summary>
    /// <remarks>
    /// Applying values the section already has is a successful no-op that does NOT consume undo
    /// history — a format dropdown firing on every selection must not evict the user's real edits
    /// from the bounded ring (same reasoning as <see cref="EnsureHeaderFooterVisible"/>).
    /// </remarks>
    public EditResult SetPageNumbering(string anchorId, PageNumberingOp op)
    {
        if (op.Start is { } s && s < 0)
            return EditResult.Fail(EditErrorCode.InvalidPageNumbering,
                "page-number start cannot be negative", anchorId);
        if (op.Format is { } f && !Internal.NumberFormats.IsPageNumberFormat(f))
            return EditResult.Fail(EditErrorCode.InvalidPageNumbering,
                $"{f} cannot format a page number", anchorId);

        return EditSectionPageNumbering(anchorId, "SetPageNumbering", sectPr =>
        {
            var existing = sectPr.Element(W.pgNumType);
            var start = op.Start?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var fmt = op.Format is { } pf ? Internal.NumberFormats.ToOoxml(pf) : null;

            if (existing is not null
                && (start is null || (string?)existing.Attribute(W.start) == start)
                && (fmt is null || (string?)existing.Attribute(W.fmt) == fmt))
                return false;
            if (existing is null && start is null && fmt is null)
                return false;

            var pgNumType = existing;
            if (pgNumType is null)
            {
                pgNumType = new XElement(W.pgNumType);
                WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, pgNumType);
            }
            if (start is not null) pgNumType.SetAttributeValue(W.start, start);
            if (fmt is not null) pgNumType.SetAttributeValue(W.fmt, fmt);
            return true;
        });
    }

    /// <summary>
    /// Remove the page-numbering setup written by <see cref="SetPageNumbering"/> from the section
    /// that owns <paramref name="anchorId"/>: the section reverts to continuing the previous
    /// section's numbering in Word's default <c>1, 2, 3</c> format.
    /// </summary>
    /// <remarks>
    /// Narrowed to <c>w:start</c> and <c>w:fmt</c> — the chapter-numbering attributes
    /// (<c>w:chapStyle</c>/<c>w:chapSep</c>), which this surface never writes, are preserved, and
    /// the <c>w:pgNumType</c> element is removed only once nothing is left on it. A section with no
    /// page numbering to clear is a successful no-op that consumes no undo history.
    /// </remarks>
    public EditResult ClearPageNumbering(string anchorId) =>
        EditSectionPageNumbering(anchorId, "ClearPageNumbering", sectPr =>
        {
            var pgNumType = sectPr.Element(W.pgNumType);
            if (pgNumType is null) return false;
            var start = pgNumType.Attribute(W.start);
            var fmt = pgNumType.Attribute(W.fmt);
            if (start is null && fmt is null) return false;

            start?.Remove();
            fmt?.Remove();
            // Only w:-namespaced attributes are CT_PageNumberType content; the element may also
            // carry pt: bookkeeping (Unid), which must not keep an otherwise-empty element alive.
            if (!pgNumType.Attributes().Any(a => a.Name.Namespace == W.w)) pgNumType.Remove();
            return true;
        });

    /// <summary>
    /// Shared body of the section page-numbering verbs: resolve a body anchor to its governing
    /// <c>w:sectPr</c> (synthesizing the document-final one if the body has none, as
    /// <see cref="SetHeaderText"/> does), then apply <paramref name="mutate"/>. The mutator returns
    /// false when the document already says what was asked, and NOTHING is snapshotted in that case.
    /// </summary>
    private EditResult EditSectionPageNumbering(string anchorId, string opName, Func<XElement, bool> mutate)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a body block anchor (the section the page numbering belongs to)",
                anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var body = element.AncestorsAndSelf(W.body).FirstOrDefault();
        if (body is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "anchor is not in the document body", anchorId);

        var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
        var succeeded = new EditResult { Success = true, Modified = new[] { target.Anchor } };

        // Decide on a detached COPY first. Returning before TakeSnapshot is what keeps a no-op out
        // of the bounded undo ring — and it also stops a no-op from synthesizing a sectPr that the
        // document did not ask for.
        if (!mutate(sectPr is null ? new XElement(W.sectPr) : new XElement(sectPr)))
            return succeeded;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (sectPr is null)
            {
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
            }
            mutate(sectPr);
            InvalidateProjectionCache();
            return succeeded;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private static void InsertSectPrTitlePg(XElement sectPr) =>
        WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, new XElement(W.titlePg));

    // ─── Footnotes / endnotes ─────────────────────────────────────────────────
    //
    // Author a note: write the definition into the FootnotesPart/EndnotesPart (creating the part
    // and Word's two reserved separator notes when it doesn't exist yet) and cite it from a body
    // paragraph at a character offset. Undo/redo of the part creation is handled by the note-part
    // reconcile in RestoreSnapshot, mirroring the header/footer one above.
    //
    // Editing and deleting an authored note need no new op: ReplaceText already accepts the note's
    // p:fn/p:en paragraph anchor, and DeleteBlock already removes an fn/en definition together with
    // every reference to it anywhere in the package (DS140/DS141). InsertFootnote/InsertEndnote
    // were the only missing verb.

    /// <summary>
    /// Create a <b>footnote</b> whose body is <paramref name="markdownPayload"/> and cite it from the
    /// paragraph named by <paramref name="anchorId"/> at <paramref name="characterOffset"/> characters
    /// into that paragraph's text (0 = before all text, text length = after all of it). Creates the
    /// <c>FootnotesPart</c>, Word's two reserved separator notes, the <c>FootnoteText</c>/
    /// <c>FootnoteReference</c> styles, and the <c>w:footnotePr</c> settings declaration if the document
    /// has no footnotes yet; otherwise the existing part is reused and only a fresh definition is added.
    /// The note id is allocated above every id already used in the package, so a document with
    /// non-contiguous note ids can't collide. Returns the created note anchors — the definition
    /// (kind <c>fn</c>) and its paragraphs (kind <c>p</c>, scope <c>fn</c>) — in
    /// <see cref="EditResult.Created"/>, so a caller can immediately edit the note with
    /// <see cref="ReplaceText"/> or remove it with <see cref="DeleteBlock"/>.
    /// </summary>
    /// <remarks>
    /// Body paragraphs only. Word does not allow a note reference inside a header/footer story or
    /// inside another note, and the projection's <c>fn</c>/<c>en</c> scopes are note <em>definitions</em>,
    /// not legal citation hosts — a non-body anchor is rejected with
    /// <see cref="EditErrorCode.AnchorWrongKind"/> rather than silently producing a document Word repairs.
    /// </remarks>
    public EditResult InsertFootnote(string anchorId, int characterOffset, string markdownPayload)
        => InsertNote(isFootnote: true, anchorId, characterOffset, markdownPayload);

    /// <summary>
    /// Create an <b>endnote</b> and cite it from a body paragraph. Behaves exactly like
    /// <see cref="InsertFootnote"/> but writes into the <c>EndnotesPart</c>, emits a
    /// <c>w:endnoteReference</c>, and uses the <c>EndnoteText</c>/<c>EndnoteReference</c> styles;
    /// the created definition anchor has kind <c>en</c> and its paragraphs scope <c>en</c>.
    /// </summary>
    public EditResult InsertEndnote(string anchorId, int characterOffset, string markdownPayload)
        => InsertNote(isFootnote: false, anchorId, characterOffset, markdownPayload);

    private EditResult InsertNote(bool isFootnote, string anchorId, int characterOffset, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var opName = isFootnote ? "InsertFootnote" : "InsertEndnote";

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a body paragraph anchor; Word does not allow a note reference in the " +
                $"'{target.Anchor.Scope}' story", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        var totalText = ParagraphText(element);
        if (characterOffset < 0 || characterOffset > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} out of [0, {totalText.Length}]", anchorId);

        // Parse the note body BEFORE snapshotting so a malformed payload is a clean no-op
        // (no part created, no undo entry pushed).
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var part = EnsureNotePart(main, isFootnote);
            Internal.StyleFactory.EnsureNoteStyles(_doc!, isFootnote);

            var noteName = isFootnote ? W.footnote : W.endnote;
            var root = part.GetXDocument().Root!;

            // Body-side citation FIRST, carrying a placeholder id — the note's real id is chosen
            // from where this citation lands among the existing ones (see NextNoteIdInReferenceOrder).
            // Split before inserting so no run or hyperlink straddles the offset: the same offset
            // mechanism SplitParagraph and ApplyFormat use, not a second walker.
            SplitRunsAtOffset(element, characterOffset);
            SplitInlineContainersAtOffset(element, characterOffset);
            var refRun = BuildNoteReferenceRun(isFootnote, NoteIdPlaceholder);
            UnidHelper.AssignToSelfAndDescendants(refRun);
            InsertInlineAtOffset(element, characterOffset, refRun);

            var id = NextNoteIdInReferenceOrder(main, root, noteName);
            refRun.Descendants(isFootnote ? W.footnoteReference : W.endnoteReference).First()
                .SetAttributeValue(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture));

            ApplyNoteBodyStyle(paras, isFootnote);
            var note = new XElement(noteName,
                new XAttribute(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                paras);
            root.Add(note);
            UnidHelper.AssignToSelfAndDescendants(note);
            part.PutXDocument();

            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var notePartUri = part.Uri.ToString();
            if (AnchorForUnid((string?)note.Attribute(PtOpenXml.Unid), notePartUri) is { } noteAnchor)
                created.Add(noteAnchor);
            foreach (var p in note.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), notePartUri) is { } pa)
                    created.Add(pa);

            return new EditResult
            {
                Success = true,
                Created = created,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Find-or-create the footnotes/endnotes part. A part created here is seeded with the two
    /// notes Word reserves for page-rendering scaffolding (<c>type="separator"</c> at id -1 and
    /// <c>type="continuationSeparator"</c> at id 0) and declared in <c>settings.xml</c>, which is
    /// what Word itself writes for a document's first note — a part holding only user notes opens,
    /// but has no separator rule above the note area.
    /// </summary>
    private static OpenXmlPart EnsureNotePart(MainDocumentPart main, bool isFootnote)
    {
        OpenXmlPart? existing = isFootnote ? main.FootnotesPart : main.EndnotesPart;
        if (existing is not null) return existing;

        var part = isFootnote ? main.AddNewPart<FootnotesPart>() : (OpenXmlPart)main.AddNewPart<EndnotesPart>();
        part.PutXDocument(new XDocument(BuildNotePartRoot(isFootnote)));
        DeclareNoteSeparatorsInSettings(main, isFootnote);
        return part;
    }

    /// <summary>A fresh <c>w:footnotes</c>/<c>w:endnotes</c> root holding only Word's two reserved notes.</summary>
    private static XElement BuildNotePartRoot(bool isFootnote)
    {
        var noteName = isFootnote ? W.footnote : W.endnote;
        // The separator marks are shared between footnotes and endnotes (there is no w:endnoteSeparator).
        XElement Reserved(string type, int id, XName mark) =>
            new XElement(noteName,
                new XAttribute(W.type, type),
                new XAttribute(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new XElement(W.p,
                    new XElement(W.pPr,
                        new XElement(W.spacing,
                            new XAttribute(W.after, "0"),
                            new XAttribute(W.line, "240"),
                            new XAttribute(W.lineRule, "auto"))),
                    new XElement(W.r, new XElement(mark))));

        return new XElement(isFootnote ? W.footnotes : W.endnotes,
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(XNamespace.Xmlns + "r", R.r),
            Reserved("separator", -1, W.separator),
            Reserved("continuationSeparator", 0, W.continuationSeparator));
    }

    /// <summary>
    /// Declare the reserved separator notes in <c>settings.xml</c> (<c>w:footnotePr</c>/
    /// <c>w:endnotePr</c>), the form Word writes. Uses the shared schema-slot insert so the
    /// settings part is never wholesale reordered (see <c>EnsureSettingsChildInOrder</c>); an
    /// existing <c>w:footnotePr</c> carrying the document's own numbering options is left alone.
    /// </summary>
    private static void DeclareNoteSeparatorsInSettings(MainDocumentPart main, bool isFootnote)
    {
        var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        var xDoc = settingsPart.GetXDocument();
        var root = xDoc.Root;
        if (root is null)
        {
            root = new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w));
            xDoc.Add(root);
        }

        var noteName = isFootnote ? W.footnote : W.endnote;
        var pr = new XElement(isFootnote ? W.footnotePr : W.endnotePr,
            new XElement(noteName, new XAttribute(W.id, "-1")),
            new XElement(noteName, new XAttribute(W.id, "0")));
        if (WordprocessingMLUtil.EnsureSettingsChildInOrder(root, pr))
            settingsPart.PutXDocument();
    }

    /// <summary>
    /// Sentinel id the citation run carries until its real id is known. Negative and far below any
    /// legal note id (Word reserves -1 and 0), so it can never be mistaken for a real citation.
    /// </summary>
    private const int NoteIdPlaceholder = int.MinValue;

    /// <summary>
    /// The id for a note whose citation (marked by <see cref="NoteIdPlaceholder"/>) is already in
    /// the body — chosen so ids stay ascending in <em>reference</em> order, shifting the notes cited
    /// after it up by one when necessary. Returns the new note's id.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Ascending-in-reference-order is an invariant every Word-authored document holds (verified
    /// across the TestFiles corpus — including documents whose ids have gaps, e.g. 17/21/26 — and
    /// the 94-footnote NVCA model certificate). Renderers rely on it: LibreOffice numbers the body
    /// markers by citation position but pairs them against the <em>id-sorted</em> definition list,
    /// so a first-cited note holding the highest id renders the WRONG note text — the marker reads
    /// "1" and points at somebody else's footnote. Nothing errors; the document is simply wrong.
    /// </para>
    /// <para>
    /// Appending <c>max(id) + 1</c> is therefore correct only when the citation follows every
    /// existing one — the common case, and the one that costs nothing here. Otherwise the new note
    /// takes the smallest id cited after it and everything at or above that shifts up, which leaves
    /// the notes cited *earlier* untouched. Taking the minimum of the following ids (rather than the
    /// first) keeps this correct even if the input document already violated the invariant.
    /// </para>
    /// </remarks>
    private static int NextNoteIdInReferenceOrder(MainDocumentPart main, XElement notePartRoot, XName noteName)
    {
        var refName = noteName == W.footnote ? W.footnoteReference : W.endnoteReference;

        // Citations in document order; the placeholder marks where the new one landed.
        var ordered = main.GetXDocument().Root!.Descendants(refName)
            .Select(r => int.TryParse((string?)r.Attribute(W.id), out var v) ? v : 0)
            .ToList();
        var following = ordered.SkipWhile(v => v != NoteIdPlaceholder).Skip(1)
            .Where(v => v != NoteIdPlaceholder).ToList();

        if (following.Count > 0)
        {
            var slot = following.Min();
            ShiftNoteIdsAtOrAbove(main, notePartRoot, noteName, slot);
            return slot;
        }

        // Cited after every existing note — appending keeps ids ascending, so no shift is needed.
        // Scan definitions AND references (a note may be cited from a header/footer or another
        // note) so a document with gaps can't alias an existing definition.
        int max = 0;
        foreach (var n in notePartRoot.Elements(noteName))
            if (int.TryParse((string?)n.Attribute(W.id), out var id) && id > max) max = id;
        foreach (var part in NoteReferenceHostParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            foreach (var r in root.Descendants(refName))
                if (int.TryParse((string?)r.Attribute(W.id), out var id) && id > max) max = id;
        }
        return max + 1;
    }

    /// <summary>
    /// Renumber every user note with id &gt;= <paramref name="threshold"/> up by one — both the
    /// definitions and every reference to them anywhere in the package — to open that id for a
    /// newly inserted note. Word-reserved notes (any <c>w:type</c>: separator,
    /// continuationSeparator, continuationNotice) are never touched; their ids sit below any user
    /// id, so shifting upward can't collide with them.
    /// </summary>
    private static void ShiftNoteIdsAtOrAbove(
        MainDocumentPart main, XElement notePartRoot, XName noteName, int threshold)
    {
        static void Bump(XElement el)
        {
            var v = int.Parse((string)el.Attribute(W.id)!, System.Globalization.CultureInfo.InvariantCulture);
            el.SetAttributeValue(W.id, (v + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        foreach (var n in notePartRoot.Elements(noteName))
        {
            if (n.Attribute(W.type) is not null) continue; // Word-reserved scaffolding
            if (int.TryParse((string?)n.Attribute(W.id), out var v) && v >= threshold) Bump(n);
        }

        var refName = noteName == W.footnote ? W.footnoteReference : W.endnoteReference;
        foreach (var part in NoteReferenceHostParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            bool any = false;
            foreach (var r in root.Descendants(refName).ToList())
                if (int.TryParse((string?)r.Attribute(W.id), out var v) && v >= threshold) { Bump(r); any = true; }
            // The main part is flushed by Save with the rest of the edit; peer parts (headers,
            // footers, the other note part) are written back here, as RemoveCrossReferences does.
            if (any && !ReferenceEquals(part, main)) part.PutXDocument();
        }
    }

    /// <summary>Every part whose XML can carry a note reference, for id-collision scanning.</summary>
    private static IEnumerable<OpenXmlPart> NoteReferenceHostParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }

    /// <summary>
    /// Shape the note body the way Word does: every paragraph gets the note-text style (unless the
    /// payload already set one) and the first paragraph opens with the auto-number mark
    /// (<c>w:footnoteRef</c>/<c>w:endnoteRef</c>) followed by a separating space, so the note reads
    /// "1 Text" rather than "1Text".
    /// </summary>
    private static void ApplyNoteBodyStyle(List<XElement> paras, bool isFootnote)
    {
        var styleId = isFootnote ? "FootnoteText" : "EndnoteText";
        foreach (var p in paras)
        {
            var pPr = p.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); p.AddFirst(pPr); }
            if (pPr.Element(W.pStyle) is null)
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));
        }

        // The loop above guarantees a w:pPr on every paragraph, so the mark + its separating
        // space go straight after the first paragraph's — schema-ordered, ahead of the payload.
        var firstPPr = paras[0].Element(W.pPr)!;
        firstPPr.AddAfterSelf(
            new XElement(W.r,
                new XElement(W.rPr,
                    new XElement(W.rStyle, new XAttribute(W.val, NoteReferenceStyleId(isFootnote)))),
                new XElement(isFootnote ? W.footnoteRef : W.endnoteRef)),
            new XElement(W.r,
                new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " ")));
    }

    /// <summary>The body-side citation run: a superscript-styled <c>w:footnoteReference</c>/
    /// <c>w:endnoteReference</c> pointing at <paramref name="id"/>.</summary>
    private static XElement BuildNoteReferenceRun(bool isFootnote, int id) =>
        new XElement(W.r,
            new XElement(W.rPr,
                new XElement(W.rStyle, new XAttribute(W.val, NoteReferenceStyleId(isFootnote)))),
            new XElement(isFootnote ? W.footnoteReference : W.endnoteReference,
                new XAttribute(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture))));

    private static string NoteReferenceStyleId(bool isFootnote) =>
        isFootnote ? "FootnoteReference" : "EndnoteReference";

    // ─── Comments (issue #300) ───────────────────────────────────────────
    //
    // Native Word comment authoring, following the part-creation pattern the note ops above
    // established: find-or-create the WordprocessingCommentsPart + the CommentText/
    // CommentReference styles, bracket a character span with w:commentRangeStart/End, append
    // the run-level w:commentReference, and add the w:comment definition. Mechanics live in
    // Internal.CommentOps; part create/delete is undo/redo-reconciled by ReconcileCommentsPart.
    //
    // Editing a comment body needs no bespoke path beyond UpdateComment: comment paragraphs
    // project as kind p, scope cmt, so ReplaceText already accepts them; DeleteBlock already
    // removes a cmt definition together with its body-side marker triple.

    /// <summary>
    /// Add a <b>native Word comment</b> (a <c>w:comment</c> the Reviewing pane shows — not the
    /// <see cref="AddAnnotation"/> overlay) on the paragraph named by <paramref name="anchorId"/>.
    /// <paramref name="span"/> selects the commented character range; <c>null</c> comments the
    /// whole block. Creates the <c>WordprocessingCommentsPart</c> and the <c>CommentText</c>/
    /// <c>CommentReference</c> styles when absent. The comment body comes from
    /// <paramref name="markdownPayload"/> (same subset as <see cref="InsertFootnote"/>).
    /// <paramref name="date"/> is written only when provided, keeping output deterministic by
    /// default; an Unspecified-kind value is treated as UTC. Returns the created definition
    /// anchor (kind <c>cmt</c>) and its paragraph anchors (kind <c>p</c>, scope <c>cmt</c>) in
    /// <see cref="EditResult.Created"/> so a caller can immediately
    /// <see cref="UpdateComment"/>/<see cref="RemoveComment"/> it.
    /// </summary>
    /// <remarks>
    /// Body paragraphs only (kind <c>p</c>/<c>h</c>/<c>li</c>, scope <c>body</c>) — Word has no
    /// comments-on-comments, and v1 does not target header/footer/note stories. Spans are
    /// single-block; the numeric <c>w:id</c> is never surfaced (comments are addressed by anchor).
    /// </remarks>
    public EditResult AddComment(
        string anchorId, CharSpan? span, string author, string markdownPayload,
        string? initials = null, DateTime? date = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"AddComment requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"AddComment requires a body paragraph anchor; got scope '{target.Anchor.Scope}'", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        var totalText = ParagraphText(element);
        int spanStart, spanLength;
        if (span.HasValue)
        {
            spanStart = span.Value.Start;
            spanLength = span.Value.Length;
            if (spanLength <= 0)
                return EditResult.Fail(EditErrorCode.EmptyCommentSpan, "span length must be > 0", anchorId);
            if (spanStart < 0 || spanStart + spanLength > totalText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{spanStart},{spanStart + spanLength}) outside block of length {totalText.Length}", anchorId);
        }
        else
        {
            spanStart = 0;
            spanLength = totalText.Length;
            if (spanLength == 0)
                return EditResult.Fail(EditErrorCode.EmptyCommentSpan, "block has no text to comment", anchorId);
        }

        // Parse the comment body BEFORE snapshotting so a malformed payload is a clean no-op
        // (no part created, no undo entry pushed).
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var part = Internal.CommentOps.EnsureCommentsPart(main);
            Internal.StyleFactory.EnsureCommentStyles(_doc!);
            var id = Internal.CommentOps.NextCommentId(main);
            var idStr = id.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // Body plumbing: bracket the span, then the reference run directly after the
            // rangeEnd — the shape Word writes. Splits route through the same offset
            // mechanism every other span op uses (AnnotationOps.SplitRunsForSpan).
            var (startRun, endRun) = Internal.AnnotationOps.SplitRunsForSpan(element, spanStart, spanLength);
            var rangeStart = new XElement(W.commentRangeStart, new XAttribute(W.id, idStr));
            var rangeEnd = new XElement(W.commentRangeEnd, new XAttribute(W.id, idStr));
            startRun.AddBeforeSelf(rangeStart);
            endRun.AddAfterSelf(rangeEnd);
            var refRun = Internal.CommentOps.BuildReferenceRun(id);
            UnidHelper.AssignToSelfAndDescendants(refRun);
            rangeEnd.AddAfterSelf(refRun);

            // Definition.
            Internal.CommentOps.ApplyCommentBodyStyle(paras);
            var comment = new XElement(W.comment,
                new XAttribute(W.id, idStr),
                new XAttribute(W.author, author));
            if (!string.IsNullOrEmpty(initials))
                comment.SetAttributeValue(W.initials, initials);
            if (date.HasValue)
                comment.SetAttributeValue(W.date, Internal.CommentOps.FormatDate(date.Value));
            foreach (var p in paras) comment.Add(p);
            var root = part.GetXDocument().Root!;
            root.Add(comment);
            UnidHelper.AssignToSelfAndDescendants(comment);
            part.PutXDocument();

            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var commentsPartUri = part.Uri.ToString();
            if (AnchorForUnid((string?)comment.Attribute(PtOpenXml.Unid), commentsPartUri) is { } defAnchor)
                created.Add(defAnchor);
            foreach (var p in comment.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), commentsPartUri) is { } pa)
                    created.Add(pa);

            return new EditResult
            {
                Success = true,
                Created = created,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Add a native Word <b>reply</b> to the comment addressed by
    /// <paramref name="parentCommentAnchorId"/>. The reply receives its own
    /// <c>w:comment</c> definition and marker id, adds an adjacent reference at the parent's
    /// native thread anchor, and links to it through <c>w15:paraIdParent</c> in a find-or-created
    /// <c>commentsExtended.xml</c>. A matching <c>commentsIds.xml</c> entry is also created for
    /// both sides when absent. Metadata ids are allocated deterministically.
    /// </summary>
    /// <remarks>
    /// The parent may itself be a reply. An orphaned definition with no live
    /// <c>w:commentReference</c> cannot be replied to because it has no document position to
    /// share. Returns the new definition and comment-body paragraph anchors in
    /// <see cref="EditResult.Created"/>.
    /// </remarks>
    public EditResult AddCommentReply(
        string parentCommentAnchorId, string author, string markdownPayload,
        string? initials = null, DateTime? date = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var parentTarget = FindAnchor(parentCommentAnchorId);
        if (parentTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"anchor not found: {parentCommentAnchorId}", parentCommentAnchorId);
        if (parentTarget.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"AddCommentReply requires a comment definition anchor (kind cmt); got kind={parentTarget.Anchor.Kind}",
                parentCommentAnchorId);

        var parentComment = parentTarget.Resolve(_doc!);
        if (parentComment is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", parentCommentAnchorId);
        var main = _doc!.MainDocumentPart;
        if (main?.WordprocessingCommentsPart is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no comments part", parentCommentAnchorId);
        var parentId = (string?)parentComment.Attribute(W.id);
        if (string.IsNullOrEmpty(parentId))
            return EditResult.Fail(EditErrorCode.InternalError,
                "parent comment definition has no w:id", parentCommentAnchorId);
        if (!Internal.CommentOps.HasCommentReference(main, parentId))
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                "parent comment definition has no live document reference", parentCommentAnchorId);

        // Parse before snapshotting so malformed markdown is a clean no-op.
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, parentCommentAnchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.StyleFactory.EnsureCommentStyles(_doc!);
            var id = Internal.CommentOps.NextCommentId(main);
            var idStr = id.ToString(System.Globalization.CultureInfo.InvariantCulture);

            // Word keeps range markers on the thread root; each reply adds only an adjacent
            // reference and inherits that range through commentsExtended parentage.
            var hostBlocks = Internal.CommentOps.InsertReplyReference(main, parentId, id);

            Internal.CommentOps.ApplyCommentBodyStyle(paras);
            var reply = new XElement(W.comment,
                new XAttribute(W.id, idStr),
                new XAttribute(W.author, author));
            if (!string.IsNullOrEmpty(initials))
                reply.SetAttributeValue(W.initials, initials);
            if (date.HasValue)
                reply.SetAttributeValue(W.date, Internal.CommentOps.FormatDate(date.Value));
            foreach (var p in paras) reply.Add(p);
            main.WordprocessingCommentsPart.GetXDocument().Root!.Add(reply);
            UnidHelper.AssignToSelfAndDescendants(reply);

            // Upgrade a flat parent only as far as needed: one extension/id entry for it and one
            // for the reply. Existing thread/resolve metadata is preserved.
            var parentParaId = Internal.CommentOps.EnsureThreadingMetadata(main, parentComment);
            Internal.CommentOps.EnsureThreadingMetadata(main, reply,
                parentParaId: parentParaId, resolved: false);

            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var commentsPartUri = main.WordprocessingCommentsPart.Uri.ToString();
            if (AnchorForUnid((string?)reply.Attribute(PtOpenXml.Unid), commentsPartUri) is { } defAnchor)
                created.Add(defAnchor);
            foreach (var p in reply.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), commentsPartUri) is { } pa)
                    created.Add(pa);

            // The parent gains/participates in extension metadata, while every document-side
            // reference host gains a new run. Report both semantic mutations; patch the first
            // host because that is the rendered document block callers need to refresh.
            var modified = new List<Anchor> { parentTarget.Anchor };
            var seenModified = new HashSet<string>(StringComparer.Ordinal)
            {
                parentTarget.Anchor.Id,
            };
            string? patchHostAnchorId = null;
            foreach (var hostBlock in hostBlocks)
            {
                if (AnchorForElement(hostBlock) is not { } anchor) continue;
                patchHostAnchorId ??= anchor.Id;
                if (seenModified.Add(anchor.Id)) modified.Add(anchor);
            }
            var patchTarget = patchHostAnchorId is null ? null : FindAnchor(patchHostAnchorId);

            return new EditResult
            {
                Success = true,
                Created = created,
                Modified = modified,
                Patch = patchTarget is null ? null : PatchFor(patchTarget),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, parentCommentAnchorId);
        }
    }

    /// <summary>
    /// Mark a comment resolved or reopened by setting <c>w15:done</c>. A flat comment is upgraded
    /// in place: its last paragraph receives a deterministic <c>w14:paraId</c>, and
    /// <c>commentsExtended.xml</c>/<c>commentsIds.xml</c> are find-or-created. Existing reply
    /// parentage is preserved. The mutation is fully undoable, including first-time part creation.
    /// </summary>
    public EditResult SetCommentResolved(string commentAnchorId, bool resolved)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(commentAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {commentAnchorId}", commentAnchorId);
        if (target.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"SetCommentResolved requires a comment definition anchor (kind cmt); got kind={target.Anchor.Kind}",
                commentAnchorId);

        var comment = target.Resolve(_doc!);
        if (comment is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", commentAnchorId);
        var main = _doc!.MainDocumentPart;
        if (main?.WordprocessingCommentsPart is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no comments part", commentAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.CommentOps.EnsureThreadingMetadata(main, comment, resolved: resolved);
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, commentAnchorId);
        }
    }

    /// <summary>
    /// Replace a comment's <b>body text</b> with <paramref name="markdownPayload"/>, addressed by
    /// its definition anchor (kind <c>cmt</c>, from <see cref="EditResult.Created"/> or the
    /// projection's <c># Comments</c> tokens). The comment's identity attributes
    /// (<c>w:id</c>/<c>w:author</c>/<c>w:initials</c>/<c>w:date</c>) are untouched. When the old
    /// last paragraph carried a <c>w14:paraId</c> (a Word-threaded comment), the id is re-stamped
    /// on the new last paragraph — <c>commentsExtended.xml</c> entries key on it, so a body edit
    /// must not orphan Word's reply/resolve metadata.
    /// </summary>
    public EditResult UpdateComment(string commentAnchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");

        var target = FindAnchor(commentAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {commentAnchorId}", commentAnchorId);
        if (target.Anchor.Kind != "cmt")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"UpdateComment requires a comment definition anchor (kind cmt); got kind={target.Anchor.Kind}",
                commentAnchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", commentAnchorId);
        var main = _doc!.MainDocumentPart;
        if (main?.WordprocessingCommentsPart is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no comments part", commentAnchorId);

        // Parse BEFORE snapshotting so a malformed payload is a clean no-op.
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, commentAnchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.StyleFactory.EnsureCommentStyles(_doc!);

            var oldParas = element.Elements(W.p).ToList();
            var preservedParaId = (string?)oldParas.LastOrDefault()?.Attribute(W14.paraId);

            // Collect the outgoing paragraph anchors before removal.
            var index = AnchorIndex();
            var removed = new List<Anchor>();
            foreach (var p in oldParas)
            {
                var unid = (string?)p.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                foreach (var kv in index)
                    if (kv.Value.Unid == unid)
                        removed.Add(kv.Value.Anchor);
            }

            foreach (var p in oldParas) p.Remove();
            Internal.CommentOps.ApplyCommentBodyStyle(paras);
            foreach (var p in paras) element.Add(p);
            if (preservedParaId is not null)
                paras[paras.Count - 1].SetAttributeValue(W14.paraId, preservedParaId);
            foreach (var p in paras) UnidHelper.AssignToSelfAndDescendants(p);
            main.WordprocessingCommentsPart.PutXDocument();

            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var commentsPartUri = main.WordprocessingCommentsPart.Uri.ToString();
            foreach (var p in element.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), commentsPartUri) is { } pa)
                    created.Add(pa);

            return new EditResult
            {
                Success = true,
                Created = created,
                Removed = removed,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RestoreSnapshot(_history.PopForUndo().snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, commentAnchorId);
        }
    }

    /// <summary>
    /// Insert <paramref name="newChild"/> into <paramref name="paragraph"/> at
    /// <paramref name="offset"/> characters into its text — before the first child that starts at
    /// or past the offset, else appended. Callers must have cleared the boundary first
    /// (<see cref="SplitRunsAtOffset"/> + <see cref="SplitInlineContainersAtOffset"/>); this is the
    /// insert-side counterpart of <see cref="MoveInlineChildrenAfter"/> and counts positions the
    /// same way, so zero-width markers sandwiched at the offset keep the ref inside their range.
    /// </summary>
    private static void InsertInlineAtOffset(XElement paragraph, int offset, XElement newChild)
    {
        int consumed = 0;
        foreach (var child in paragraph.Elements().ToList())
        {
            if (child.Name == W.pPr) continue;
            if (consumed >= offset) { child.AddBeforeSelf(newChild); return; }
            consumed += IsInlineChild(child) ? InlineChildTextLength(child) : 0;
        }
        paragraph.Add(newChild);
    }

    /// <summary>
    /// Insert a <paramref name="rows"/>×<paramref name="cols"/> table before/after the block named
    /// by <paramref name="anchorId"/>. <paramref name="options"/> controls borders, per-cell markdown
    /// (row-major), and cell alignment. Returns the created cell-paragraph anchors (row-major), so the
    /// caller can address and fill/format each cell.
    /// </summary>
    public EditResult InsertTable(string anchorId, Position pos, int rows, int cols, TableInsertOptions? options = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (rows < 1 || cols < 1)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "table needs >= 1 row and >= 1 column", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        var opts = options ?? new TableInsertOptions();
        var contents = opts.CellContents;

        // Explicit per-column widths: one per column, all positive. A mismatched count is a
        // caller error — reject rather than silently equalize (no silent caps).
        var colWidths = opts.ColumnWidths;
        if (colWidths is not null && (colWidths.Count != cols || colWidths.Any(w => w <= 0)))
            return EditResult.Fail(EditErrorCode.MalformedMarkdown,
                $"ColumnWidths must have one positive width per column ({cols}); got {colWidths.Count}", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            const int contentTwips = 9576;           // ~6.65", a US-Letter content width
            int colTwips = contentTwips / cols;
            int Width(int c) => colWidths is not null ? colWidths[c] : colTwips;

            // With explicit widths the table is sized to their sum (dxa); otherwise it fills
            // the content area (100% pct) and splits equally.
            var tblW = colWidths is not null
                ? new XElement(W.tblW, new XAttribute(W._w, colWidths.Sum()), new XAttribute(W.type, "dxa"))
                : new XElement(W.tblW, new XAttribute(W._w, 5000), new XAttribute(W.type, "pct"));

            var tblPr = new XElement(W.tblPr,
                tblW,
                BuildTableBorders(opts.Borderless),
                new XElement(W.tblLayout, new XAttribute(W.type, "fixed")));

            var tblGrid = new XElement(W.tblGrid);
            for (int c = 0; c < cols; c++)
                tblGrid.Add(new XElement(W.gridCol, new XAttribute(W._w, Width(c))));

            var tbl = new XElement(W.tbl, tblPr, tblGrid);
            var cellParagraphs = new List<XElement>();

            for (int r = 0; r < rows; r++)
            {
                var tr = new XElement(W.tr);
                for (int c = 0; c < cols; c++)
                {
                    var tc = new XElement(W.tc,
                        new XElement(W.tcPr, new XElement(W.tcW, new XAttribute(W._w, Width(c)), new XAttribute(W.type, "dxa"))));

                    int idx = r * cols + c;
                    string? md = contents is not null && idx < contents.Count ? contents[idx] : null;
                    var paras = BuildCellParagraphs(md, opts.CellAlignment);
                    foreach (var p in paras) tc.Add(p);
                    cellParagraphs.AddRange(paras);
                    tr.Add(tc);
                }
                tbl.Add(tr);
            }

            UnidHelper.AssignToSelfAndDescendants(tbl);

            if (pos == Position.Before) element.AddBeforeSelf(tbl);
            else element.AddAfterSelf(tbl);

            // A table must be followed by a paragraph: Word's convention is to keep a w:p after
            // every table, and an end-of-body table with no trailing paragraph leaves no editable
            // block below it (S-1 smoke-test finding 2). If nothing — or only a sectPr / another
            // table — follows, append an empty trailing paragraph.
            var afterTbl = tbl.ElementsAfterSelf().FirstOrDefault();
            if (afterTbl is null || afterTbl.Name == W.sectPr || afterTbl.Name == W.tbl)
            {
                var trailing = new XElement(W.p);
                UnidHelper.AssignToSelfAndDescendants(trailing);
                tbl.AddAfterSelf(trailing);
            }

            foreach (var p in cellParagraphs) PromoteHyperlinkRelationships(p);

            InvalidateProjectionCache();
            var index = AnchorIndex();
            var created = new List<Anchor>();
            foreach (var p in cellParagraphs)
            {
                var unid = (string)p.Attribute(PtOpenXml.Unid)!;
                if (AnchorForUnid(unid, PartUriOf(p)) is { } a)
                    created.Add(a);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>Build the cell's paragraph(s) from optional markdown + alignment. Always >= 1 paragraph.</summary>
    private static List<XElement> BuildCellParagraphs(string? markdown, ParagraphAlignment? align)
    {
        var result = new List<XElement>();
        if (!string.IsNullOrEmpty(markdown))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdown);
            if (parsed.Success)
                foreach (var block in parsed.Blocks)
                    result.Add(BuildParagraphFromParsedBlock(block));
        }
        if (result.Count == 0) result.Add(new XElement(W.p));

        if (align is { } a)
        {
            var val = a switch
            {
                ParagraphAlignment.Center => "center",
                ParagraphAlignment.Right => "right",
                ParagraphAlignment.Justify => "both",
                _ => "left",
            };
            foreach (var p in result)
            {
                var pPr = p.Element(W.pPr);
                if (pPr is null) { pPr = new XElement(W.pPr); p.AddFirst(pPr); }
                SetPPrChildInOrder(pPr, new XElement(W.jc, new XAttribute(W.val, val)));
            }
        }
        return result;
    }

    private static XElement BuildTableBorders(bool borderless)
    {
        var edges = new[] { W.top, W.left, W.bottom, W.right, W.insideH, W.insideV };
        var bdr = new XElement(W.tblBorders);
        foreach (var e in edges)
            bdr.Add(borderless
                ? new XElement(e, new XAttribute(W.val, "none"), new XAttribute(W.sz, 0),
                    new XAttribute(W.space, 0), new XAttribute(W.color, "auto"))
                : new XElement(e, new XAttribute(W.val, "single"), new XAttribute(W.sz, 4),
                    new XAttribute(W.space, 0), new XAttribute(W.color, "auto")));
        return bdr;
    }

    // ─── Table editing (row / column CRUD), addressed by a cell-paragraph anchor ──────────
    //
    // v1 assumes a rectangular grid with no horizontal cell merges (w:gridSpan) — the shape
    // InsertTable produces and the common case for layout tables (the S-1 columns).

    /// <summary>Resolve a cell-paragraph anchor to its (paragraph, cell, row, table, column index,
    /// anchor target). Returns a failure EditResult via <paramref name="error"/> on any miss.</summary>
    private EditResult? ResolveCell(string cellAnchorId, out XElement? p, out XElement? tc,
        out XElement? tr, out XElement? tbl, out int colIndex, out AnchorTarget? target)
    {
        p = tc = tr = tbl = null; colIndex = -1; target = null;
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        target = FindAnchor(cellAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {cellAnchorId}", cellAnchorId);
        p = target.Resolve(_doc!);
        if (p is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", cellAnchorId);
        tc = p.Ancestors(W.tc).FirstOrDefault();
        if (tc is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "table row/column ops require an anchor inside a table cell", cellAnchorId);
        tr = tc.Ancestors(W.tr).FirstOrDefault();
        tbl = tr?.Ancestors(W.tbl).FirstOrDefault();
        if (tr is null || tbl is null)
            return EditResult.Fail(EditErrorCode.InternalError, "malformed table (cell has no row/table)", cellAnchorId);
        colIndex = tr.Elements(W.tc).ToList().IndexOf(tc);
        return null;
    }

    /// <summary>After a structural edit, resolve the freshly-projected anchors for the given paragraphs.</summary>
    private List<Anchor> ResolveAnchorsForParagraphs(IEnumerable<XElement> paras)
    {
        var index = AnchorIndex();
        var result = new List<Anchor>();
        foreach (var para in paras)
        {
            var unid = (string?)para.Attribute(PtOpenXml.Unid);
            if (unid is not null && AnchorForUnid(unid, PartUriOf(para)) is { } a)
                result.Add(a);
        }
        return result;
    }

    private static XElement NewEmptyCellLike(XElement referenceCell)
    {
        var tcPr = referenceCell.Element(W.tcPr);
        var tc = new XElement(W.tc);
        if (tcPr is not null) tc.Add(new XElement(tcPr)); // clone width/borders/valign
        var p = new XElement(W.p);
        tc.Add(p);
        return tc;
    }

    /// <summary>Insert a row before/after the row containing <paramref name="cellAnchorId"/>. The new
    /// row clones each column's cell width and starts empty. Returns the new cell-paragraph anchors.</summary>
    public EditResult InsertTableRow(string cellAnchorId, Position pos)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out var tr, out _, out _, out var target) is { } err)
            return err;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var newTr = new XElement(W.tr);
            var newParas = new List<XElement>();
            foreach (var tc in tr!.Elements(W.tc))
            {
                var newTc = NewEmptyCellLike(tc);
                newParas.Add(newTc.Element(W.p)!);
                newTr.Add(newTc);
            }
            UnidHelper.AssignToSelfAndDescendants(newTr);
            if (pos == Position.Before) tr.AddBeforeSelf(newTr);
            else tr.AddAfterSelf(newTr);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = ResolveAnchorsForParagraphs(newParas),
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>Insert a column before/after the column containing <paramref name="cellAnchorId"/>: a new
    /// cell in every row (cloning that column's width) plus a matching w:gridCol. Returns the new
    /// cell-paragraph anchors (top→bottom).</summary>
    public EditResult InsertTableColumn(string cellAnchorId, Position pos)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out _, out var tbl, out var colIndex, out var target) is { } err)
            return err;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var newParas = new List<XElement>();
            foreach (var tr in tbl!.Elements(W.tr))
            {
                var cells = tr.Elements(W.tc).ToList();
                var refTc = colIndex < cells.Count ? cells[colIndex] : cells[^1];
                var newTc = NewEmptyCellLike(refTc);
                UnidHelper.AssignToSelfAndDescendants(newTc);
                newParas.Add(newTc.Element(W.p)!);
                if (pos == Position.Before) refTc.AddBeforeSelf(newTc);
                else refTc.AddAfterSelf(newTc);
            }

            // Mirror the structural change in w:tblGrid so column count stays consistent.
            var grid = tbl.Element(W.tblGrid);
            if (grid is not null)
            {
                var cols = grid.Elements(W.gridCol).ToList();
                if (colIndex < cols.Count)
                {
                    var clone = new XElement(cols[colIndex]);
                    if (pos == Position.Before) cols[colIndex].AddBeforeSelf(clone);
                    else cols[colIndex].AddAfterSelf(clone);
                }
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = ResolveAnchorsForParagraphs(newParas),
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>Delete the row containing <paramref name="cellAnchorId"/>. Deleting the last row removes
    /// the whole table.</summary>
    public EditResult DeleteTableRow(string cellAnchorId)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out var tr, out var tbl, out _, out var target) is { } err)
            return err;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var index = AnchorIndex();
            var removed = CellParagraphAnchorsIn(tr!);
            if (tbl!.Elements(W.tr).Count() <= 1) { foreach (var a in CellParagraphAnchorsIn(tbl)) if (!removed.Contains(a)) removed.Add(a); tbl.Remove(); }
            else tr!.Remove();

            InvalidateProjectionCache();
            return new EditResult { Success = true, Removed = removed, Patch = PatchFor(target!) };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>Delete the column containing <paramref name="cellAnchorId"/> from every row (and its
    /// w:gridCol). Deleting the last column removes the whole table.</summary>
    public EditResult DeleteTableColumn(string cellAnchorId)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out _, out var tbl, out var colIndex, out var target) is { } err)
            return err;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var index = AnchorIndex();
            var grid = tbl!.Element(W.tblGrid);
            int colCount = grid?.Elements(W.gridCol).Count() ?? tbl.Elements(W.tr).First().Elements(W.tc).Count();

            var removed = new List<Anchor>();
            if (colCount <= 1) { foreach (var a in CellParagraphAnchorsIn(tbl)) removed.Add(a); tbl.Remove(); }
            else
            {
                foreach (var tr in tbl.Elements(W.tr).ToList())
                {
                    var cells = tr.Elements(W.tc).ToList();
                    if (colIndex >= cells.Count) continue;
                    foreach (var a in CellParagraphAnchorsIn(cells[colIndex])) removed.Add(a);
                    cells[colIndex].Remove();
                }
                var cols = grid?.Elements(W.gridCol).ToList();
                if (cols is not null && colIndex < cols.Count) cols[colIndex].Remove();
            }

            InvalidateProjectionCache();
            return new EditResult { Success = true, Removed = removed, Patch = PatchFor(target!) };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>The cell-paragraph anchors under <paramref name="scope"/> (a tc/tr/tbl), in document order.
    /// Resolved via <see cref="AnchorForElement"/> so a table inside a header/footer story cannot
    /// report a body paragraph's anchor through a colliding content-addressed unid.</summary>
    private List<Anchor> CellParagraphAnchorsIn(XElement scope)
    {
        var result = new List<Anchor>();
        foreach (var para in scope.Descendants(W.p))
        {
            if (AnchorForElement(para) is { } a) result.Add(a);
        }
        return result;
    }

    // ─── Table styling (issue #315 Stage A), addressed by a cell-paragraph anchor ─────────
    //
    // Localized w:tblPr / w:trPr / w:tcPr writes over the same rectangular-grid v1 model the
    // row/column CRUD assumes. Cell merge (w:gridSpan/w:vMerge) is Stage B and needs its own
    // design pass first.

    // CT_TblPr / CT_TcPr / CT_TrPr / CT_TblBorders child schema order (local names), matching
    // WordprocessingMLUtil's ordering tables.
    private static readonly string[] TblPrChildOrder =
    {
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
        "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
        "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
    };

    private static readonly string[] TcPrChildOrder =
    {
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
        "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers",
    };

    private static readonly string[] TrPrChildOrder =
    {
        "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
        "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden",
    };

    private static readonly string[] TblBordersEdgeOrder =
    {
        "top", "left", "start", "bottom", "right", "end", "insideH", "insideV",
    };

    /// <summary>Insert (replacing any existing) a child at its correct schema position per
    /// <paramref name="order"/> — the generalized <see cref="SetPPrChildInOrder"/>.</summary>
    private static void SetChildInOrder(XElement parent, XElement child, string[] order)
    {
        parent.Elements(child.Name).Remove();
        int idx = Array.IndexOf(order, child.Name.LocalName);
        XElement? after = null;
        foreach (var e in parent.Elements())
        {
            int ei = Array.IndexOf(order, e.Name.LocalName);
            if (ei >= 0 && ei < idx) after = e;
            else if (ei >= idx) break;
        }
        if (after is null) parent.AddFirst(child);
        else after.AddAfterSelf(child);
    }

    /// <summary>w:tblPr must be the table's first child.</summary>
    private static XElement GetOrCreateTblPr(XElement tbl)
    {
        var tblPr = tbl.Element(W.tblPr);
        if (tblPr is null) { tblPr = new XElement(W.tblPr); tbl.AddFirst(tblPr); }
        return tblPr;
    }

    /// <summary>w:tcPr must be the cell's first child.</summary>
    private static XElement GetOrCreateTcPr(XElement tc)
    {
        var tcPr = tc.Element(W.tcPr);
        if (tcPr is null) { tcPr = new XElement(W.tcPr); tc.AddFirst(tcPr); }
        return tcPr;
    }

    /// <summary>The shared "styling applied" result: the target anchor in Modified + a patch.</summary>
    private EditResult TableStyleResult(AnchorTarget target)
    {
        InvalidateProjectionCache();
        var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;
        return new EditResult
        {
            Success = true,
            Modified = new[] { updated },
            Patch = PatchFor(target),
        };
    }

    /// <summary>
    /// Retune the column widths of the table containing <paramref name="cellAnchorId"/> —
    /// the post-insert counterpart of <see cref="TableInsertOptions.ColumnWidths"/>. Rewrites
    /// <c>w:tblGrid</c> and every row's <c>w:tcW</c>, sizes the table to the widths' sum
    /// (dxa) and pins <c>w:tblLayout</c> fixed, exactly as inserting with explicit widths
    /// would. One positive twip value per column is required.
    /// </summary>
    public EditResult SetColumnWidths(string cellAnchorId, IReadOnlyList<int> widthsTwips)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out _, out var tbl, out _, out var target) is { } err)
            return err;

        var grid = tbl!.Element(W.tblGrid);
        int colCount = grid is not null
            ? grid.Elements(W.gridCol).Count()
            : tbl.Elements(W.tr).First().Elements(W.tc).Count();
        if (widthsTwips is null || widthsTwips.Count != colCount || widthsTwips.Any(w => w <= 0))
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                $"widths must list one positive twip value per column ({colCount}); got {widthsTwips?.Count ?? 0}",
                cellAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (grid is null)
            {
                grid = new XElement(W.tblGrid);
                var pr = tbl.Element(W.tblPr);
                if (pr is not null) pr.AddAfterSelf(grid);
                else tbl.AddFirst(grid);
            }
            grid.RemoveNodes();
            foreach (var w in widthsTwips)
                grid.Add(new XElement(W.gridCol, new XAttribute(W._w, w)));

            foreach (var tr in tbl.Elements(W.tr))
            {
                var cells = tr.Elements(W.tc).ToList();
                for (int c = 0; c < cells.Count && c < colCount; c++)
                    SetChildInOrder(GetOrCreateTcPr(cells[c]),
                        new XElement(W.tcW, new XAttribute(W._w, widthsTwips[c]), new XAttribute(W.type, "dxa")),
                        TcPrChildOrder);
            }

            var tblPr = GetOrCreateTblPr(tbl);
            SetChildInOrder(tblPr,
                new XElement(W.tblW, new XAttribute(W._w, widthsTwips.Sum()), new XAttribute(W.type, "dxa")),
                TblPrChildOrder);
            SetChildInOrder(tblPr,
                new XElement(W.tblLayout, new XAttribute(W.type, "fixed")),
                TblPrChildOrder);

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>
    /// Set the table-level borders (<c>w:tblPr/w:tblBorders</c>) of the table containing
    /// <paramref name="cellAnchorId"/>. Only the edges named by <see cref="TableBorderSpec.Scope"/>
    /// are written (as explicit edges, so style-inherited borders are overridden); the rest are
    /// left untouched. Style "none" removes the targeted edges the way
    /// <see cref="TableInsertOptions.Borderless"/> does. Cell-level <c>w:tcBorders</c>, where a
    /// document has them, still win over these — v1 does not touch per-cell borders.
    /// </summary>
    public EditResult SetTableBorders(string cellAnchorId, TableBorderSpec? spec = null)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out _, out var tbl, out _, out var target) is { } err)
            return err;

        var s = spec ?? new TableBorderSpec();
        if (s.Size is < 0)
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                "border size (eighths of a point) must be >= 0", cellAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var tblPr = GetOrCreateTblPr(tbl!);
            var borders = tblPr.Element(W.tblBorders);
            if (borders is null)
            {
                borders = new XElement(W.tblBorders);
                SetChildInOrder(tblPr, borders, TblPrChildOrder);
            }

            var edges = s.Scope switch
            {
                TableBorderScope.Outside => new[] { W.top, W.left, W.bottom, W.right },
                TableBorderScope.Inside => new[] { W.insideH, W.insideV },
                _ => new[] { W.top, W.left, W.bottom, W.right, W.insideH, W.insideV },
            };

            bool none = string.Equals(s.Style, "none", StringComparison.OrdinalIgnoreCase);
            foreach (var edgeName in edges)
            {
                var edge = none
                    ? new XElement(edgeName, new XAttribute(W.val, "none"), new XAttribute(W.sz, 0),
                        new XAttribute(W.space, 0), new XAttribute(W.color, "auto"))
                    : new XElement(edgeName,
                        new XAttribute(W.val, string.IsNullOrEmpty(s.Style) ? "single" : s.Style),
                        new XAttribute(W.sz, s.Size ?? 4),
                        new XAttribute(W.space, 0),
                        new XAttribute(W.color, string.IsNullOrEmpty(s.Color) ? "auto" : s.Color));
                SetChildInOrder(borders, edge, TblBordersEdgeOrder);
            }

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>
    /// Shade the cell containing <paramref name="cellAnchorId"/> — or, with
    /// <see cref="TableShadingScope.Row"/>, every cell of its row (header-row banding).
    /// <paramref name="fillColor"/> is a hex RRGGBB triplet (a leading '#' is tolerated) or
    /// "auto"; null/empty removes the shading. Writes <c>w:tcPr/w:shd</c> with
    /// <c>w:val="clear"</c>, Word's plain-fill idiom.
    /// </summary>
    public EditResult SetCellShading(string cellAnchorId, string? fillColor,
        TableShadingScope scope = TableShadingScope.Cell)
    {
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out _, out _, out var target) is { } err)
            return err;

        bool clear = string.IsNullOrEmpty(fillColor);
        string fill = "auto";
        if (!clear)
        {
            fill = fillColor!.TrimStart('#');
            if (!string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(fill, "^[0-9A-Fa-f]{6}$"))
                    return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                        $"fill must be a hex RRGGBB triplet or \"auto\"; got '{fillColor}'", cellAnchorId);
                fill = fill.ToUpperInvariant();
            }
            else fill = "auto";
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var cells = scope == TableShadingScope.Row ? tr!.Elements(W.tc).ToList() : new List<XElement> { tc! };
            foreach (var cell in cells)
            {
                if (clear)
                {
                    cell.Element(W.tcPr)?.Elements(W.shd).Remove();
                    continue;
                }
                SetChildInOrder(GetOrCreateTcPr(cell),
                    new XElement(W.shd, new XAttribute(W.val, "clear"), new XAttribute(W.color, "auto"),
                        new XAttribute(W.fill, fill)),
                    TcPrChildOrder);
            }

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>
    /// Mark (or unmark) the row containing <paramref name="cellAnchorId"/> as a repeating
    /// header row (<c>w:trPr/w:tblHeader</c>), so a multi-page table re-shows it on every page.
    /// Word only honors the flag on a run of rows starting at the table's first row — setting
    /// it elsewhere is legal but ignored by renderers.
    /// </summary>
    public EditResult SetRepeatHeaderRow(string cellAnchorId, bool repeat)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out var tr, out _, out _, out var target) is { } err)
            return err;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var trPr = tr!.Element(W.trPr);
            if (repeat)
            {
                if (trPr is null) { trPr = new XElement(W.trPr); tr.AddFirst(trPr); }
                SetChildInOrder(trPr, new XElement(W.tblHeader), TrPrChildOrder);
            }
            else if (trPr is not null)
            {
                trPr.Elements(W.tblHeader).Remove();
                // An emptied trPr is dropped entirely. Only element children matter: CT_TrPr has
                // no schema attributes, and the in-memory tree may carry pt bookkeeping attributes
                // (Unid) that Save() strips anyway.
                if (!trPr.HasElements) trPr.Remove();
            }

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    public EditResult SetListLevel(string anchorId, int levelDelta)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetListLevel requires a list-item anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var pPr = element.Element(W.pPr);
        var numPr = pPr?.Element(W.numPr);

        // Resolve the effective (numId, current ilvl). A direct w:numPr wins; otherwise the
        // paragraph is a list item only via its pStyle chain (e.g. python-docx "List Bullet",
        // which carries numPr on the STYLE, not the paragraph). In that case read the effective
        // values from the style and materialize a direct w:numPr below — exactly what Word does
        // when you Tab a styled list item, and the only way to control ilvl per paragraph.
        int current;
        int? effectiveNumId;
        if (numPr is not null)
        {
            current = (int?)numPr.Element(W.ilvl)?.Attribute(W.val) ?? 0;
            effectiveNumId = (int?)numPr.Element(W.numId)?.Attribute(W.val);
        }
        else
        {
            (effectiveNumId, current) = ResolveStyleNumbering(element);
            if (effectiveNumId is null)
                return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                    "no numPr on this paragraph or its style", anchorId);
        }

        int next = current + levelDelta;
        if (next < 0 || next > 8)
            return EditResult.Fail(EditErrorCode.InvalidListLevel,
                $"resulting list level {next} out of [0,8]", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        // Nesting only renders if the abstractNum actually DEFINES the target level — many docs
        // define just level 0, so synthesize any missing levels before bumping ilvl.
        if (effectiveNumId.HasValue)
            Internal.NumberingFactory.EnsureLevelDefined(_doc!, effectiveNumId.Value, next);

        if (numPr is not null)
        {
            numPr.Element(W.ilvl)?.Remove();
            numPr.AddFirst(new XElement(W.ilvl, new XAttribute(W.val, next))); // ilvl precedes numId
        }
        else
        {
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            SetPPrChildInOrder(pPr, new XElement(W.numPr,
                new XElement(W.ilvl, new XAttribute(W.val, next)),
                new XElement(W.numId, new XAttribute(W.val, effectiveNumId!.Value))));
        }
        // Flush the body mutation to the part stream immediately — same as NumberingFactory does for
        // the numbering part. Without this the materialized w:numPr lives only in the in-memory
        // XDocument; under WASM the typed-DOM/XDocument divergence means a later Save() serializes
        // the un-flushed state and the nest silently vanishes on save and re-render. (Body lists are
        // body-scoped; flushing the main part covers them.)
        _doc!.MainDocumentPart!.PutXDocument();
        InvalidateProjectionCache();
        return new EditResult
        {
            Success = true,
            Modified = new[] { target.Anchor },
            Patch = PatchFor(target),
        };
    }

    /// <summary>
    /// Resolve the effective <c>(numId, ilvl)</c> a paragraph inherits from its pStyle chain, for
    /// a list item whose numbering comes from a style rather than a direct <c>w:numPr</c>. Walks
    /// <c>basedOn</c> (cycle-guarded). Returns <c>(null, 0)</c> when no style contributes a numId.
    /// </summary>
    private (int? numId, int ilvl) ResolveStyleNumbering(XElement paragraph)
    {
        var styleId = (string?)paragraph.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (string.IsNullOrEmpty(styleId)) return (null, 0);
        var stylesRoot = _doc!.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
        if (stylesRoot is null) return (null, 0);

        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        for (int i = 0; i < 16 && current is not null; i++)
        {
            if (!visited.Add(current)) break; // cycle
            var style = stylesRoot.Elements(W.style)
                .FirstOrDefault(s => (string?)s.Attribute(W.styleId) == current);
            if (style is null) break;
            var styleNumPr = style.Element(W.pPr)?.Element(W.numPr);
            var numId = (int?)styleNumPr?.Element(W.numId)?.Attribute(W.val);
            if (numId is not null)
                return (numId, (int?)styleNumPr!.Element(W.ilvl)?.Attribute(W.val) ?? 0);
            current = (string?)style.Element(W.basedOn)?.Attribute(W.val);
        }
        return (null, 0);
    }

    public EditResult RemoveListMembership(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "RemoveListMembership requires list-item anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        element.Element(W.pPr)?.Element(W.numPr)?.Remove();
        InvalidateProjectionCache();
        var fresh = AnchorIndex();
        var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;
        return new EditResult
        {
            Success = true,
            Modified = new[] { updated },
            Patch = PatchFor(target),
        };
    }

    /// <summary>
    /// Make the paragraph a bullet or numbered list item, or remove list membership.
    /// Unlike <see cref="SetListLevel"/>/<see cref="RemoveListMembership"/> (which require an
    /// existing list item), this PROMOTES a plain paragraph: it ensures a reusable numbering
    /// definition exists (synthesizing one in the numbering part if needed) and sets the
    /// paragraph's <c>w:numPr</c>. <see cref="ListFormat.None"/> strips inline list membership.
    /// </summary>
    public EditResult ApplyListFormat(string anchorId, ListFormat kind)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ApplyListFormat requires a paragraph anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (kind == ListFormat.None)
            {
                pPr?.Element(W.numPr)?.Remove();
            }
            else
            {
                if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
                int numId = Internal.NumberingFactory.EnsureNumbering(_doc!, kind);
                int ilvl = (int?)pPr.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
                pPr.Element(W.numPr)?.Remove();
                SetPPrChildInOrder(pPr, new XElement(W.numPr,
                    new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
                    new XElement(W.numId, new XAttribute(W.val, numId))));
            }

            InvalidateProjectionCache();
            var freshIndex = AnchorIndex();
            var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;
            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// <see cref="ApplyListFormat"/> across a contiguous sibling run of paragraphs, from
    /// <paramref name="firstAnchorId"/> to <paramref name="lastAnchorId"/> INCLUSIVE (they may
    /// be the same anchor, and may be passed in either document order). Every member gets the
    /// same shared <c>w:num</c> instance, so the numbering sequence stays intact — the per-item
    /// op cannot guarantee that. One snapshot is recorded, so the whole range is a single
    /// <see cref="Undo"/> step. Non-paragraph siblings inside the range (a table, an sdt) are
    /// skipped — they cannot carry <c>w:numPr</c>. Each paragraph keeps its own <c>w:ilvl</c>,
    /// so a nested run converts in place. <see cref="ListFormat.None"/> strips inline list
    /// membership from every member.
    /// </summary>
    public EditResult ApplyListFormatRange(string firstAnchorId, string lastAnchorId, ListFormat kind)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var firstTarget = FindAnchor(firstAnchorId);
        if (firstTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"first anchor not found: {firstAnchorId}", firstAnchorId);
        var lastTarget = FindAnchor(lastAnchorId);
        if (lastTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"last anchor not found: {lastAnchorId}", lastAnchorId);
        if (firstTarget.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ApplyListFormatRange requires paragraph anchors; first kind={firstTarget.Anchor.Kind}", firstAnchorId);
        if (lastTarget.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ApplyListFormatRange requires paragraph anchors; last kind={lastTarget.Anchor.Kind}", lastAnchorId);
        if (firstTarget.Anchor.Scope != lastTarget.Anchor.Scope)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                $"ApplyListFormatRange anchors must live in the same package part; first={firstTarget.Anchor.Scope} last={lastTarget.Anchor.Scope}",
                firstAnchorId);

        var firstElement = firstTarget.Resolve(_doc!);
        var lastElement = lastTarget.Resolve(_doc!);
        if (firstElement is null || lastElement is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", firstAnchorId);
        if (firstElement.Parent != lastElement.Parent)
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "ApplyListFormatRange anchors must share a direct parent (no spanning into nested containers)",
                firstAnchorId);

        // Normalize order: same parent is established, so if `last` is not a following sibling
        // of `first`, the caller passed them reversed — swap rather than erroring.
        if (firstElement != lastElement && !firstElement.ElementsAfterSelf().Contains(lastElement))
            (firstElement, lastElement) = (lastElement, firstElement);

        // The w:p members of the run, first..last inclusive, with their unids captured pre-op
        // so the post-op anchors (kind may flip p↔li) can be reported in Modified.
        var members = new List<XElement>();
        for (var cursor = firstElement; cursor is not null; cursor = cursor.ElementsAfterSelf().FirstOrDefault())
        {
            if (cursor.Name == W.p) members.Add(cursor);
            if (cursor == lastElement) break;
        }
        var memberUnids = members.Select(m => (string?)m.Attribute(PtOpenXml.Unid)).ToList();
        var partUri = firstTarget.PartUri;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (kind == ListFormat.None)
            {
                foreach (var member in members)
                    member.Element(W.pPr)?.Element(W.numPr)?.Remove();
            }
            else
            {
                // One find-or-create up front — every member points at the SAME numId.
                int numId = Internal.NumberingFactory.EnsureNumbering(_doc!, kind);
                foreach (var member in members)
                {
                    var pPr = member.Element(W.pPr);
                    if (pPr is null) { pPr = new XElement(W.pPr); member.AddFirst(pPr); }
                    int ilvl = (int?)pPr.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
                    pPr.Element(W.numPr)?.Remove();
                    SetPPrChildInOrder(pPr, new XElement(W.numPr,
                        new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
                        new XElement(W.numId, new XAttribute(W.val, numId))));
                }
            }

            InvalidateProjectionCache();
            _ = AnchorIndex();
            var modified = new List<Anchor>();
            foreach (var unid in memberUnids)
            {
                if (unid is not null && AnchorForUnid(unid, partUri) is { } updated)
                    modified.Add(updated);
            }
            return new EditResult
            {
                Success = true,
                Modified = modified,
                Patch = PatchFor(firstTarget),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, firstAnchorId);
        }
    }

    /// <summary>
    /// Restart (or seed) the anchored list item's numbering at <paramref name="value"/> — Word's
    /// <em>Set Numbering Value… → Set value to</em> (issue #314). Writes a
    /// <c>w:lvlOverride[@w:ilvl]/w:startOverride[@w:val]</c> on a DEDICATED <c>w:num</c> instance:
    /// the item's current num is cloned (never mutated — it may be shared, and the numbering part
    /// is not snapshotted for undo), and the anchored paragraph plus every FOLLOWING paragraph of
    /// the same numbering instance in the part is repointed at the clone. An anchored item
    /// mid-sequence therefore splits the sequence exactly like Word: earlier items keep their
    /// numbers, the anchored item shows <paramref name="value"/>, and the tail continues from it.
    /// Style-derived members get a direct <c>w:numPr</c> materialized (ilvl preserved), the same
    /// way <see cref="SetListLevel"/> does. Undo restores every repointed paragraph.
    /// </summary>
    public EditResult SetListStartOverride(string anchorId, int value)
    {
        if (value < 0)
            return EditResult.Fail(EditErrorCode.InvalidListStartValue,
                $"list start value cannot be negative (got {value})", anchorId);
        return ApplyListStartOverride(anchorId, value);
    }

    /// <summary>
    /// Remove the numbering restart from the anchored item's list sequence — the inverse of
    /// <see cref="SetListStartOverride"/>. EVERY paragraph of the same numbering instance in the
    /// part (before and after the anchor — they move together, so relative continuation is
    /// preserved) is repointed at a clone of the instance WITHOUT the
    /// <c>w:startOverride</c> at the item's level; the sequence reverts to the abstract
    /// definition's own <c>w:start</c>. A sequence with no override at the item's level is a
    /// successful no-op that consumes no undo history.
    /// </summary>
    public EditResult ClearListStartOverride(string anchorId) =>
        ApplyListStartOverride(anchorId, null);

    /// <summary>Shared engine for <see cref="SetListStartOverride"/> (split the sequence at the
    /// anchor onto a clone carrying the override) and <see cref="ClearListStartOverride"/>
    /// (move the whole sequence onto a clone without it).</summary>
    private EditResult ApplyListStartOverride(string anchorId, int? value)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "SetListStartOverride requires a list-item anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var (numId, ilvl) = EffectiveNumberingOf(element);
        // numId 0 is OOXML for "numbering removed" — not an instance a start override can target.
        if (numId is null or 0)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "no numPr on this paragraph or its style", anchorId);

        // Clearing a sequence that has no override at this level is a no-op; return BEFORE
        // TakeSnapshot so it cannot evict real edits from the bounded undo ring.
        if (value is null && Internal.NumberingFactory.GetStartOverride(_doc!, numId.Value, ilvl) is null)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var newNumId = Internal.NumberingFactory.CloneNumWithStartOverride(_doc!, numId.Value, ilvl, value);
            if (newNumId is null)
            {
                _ = _history.PopForUndo();
                return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                    $"numbering instance {numId} is not defined in the numbering part", anchorId);
            }

            // Repoint the sequence: for a set, the anchored paragraph and everything after it
            // (the split Word performs); for a clear, every member (the whole sequence moves).
            var partRoot = element.AncestorsAndSelf().Last();
            var repointedUnids = new List<string?>();
            bool reached = false;
            foreach (var p in partRoot.Descendants(W.p))
            {
                if (p == element) reached = true;
                if (value is not null && !reached) continue;
                var (pNumId, pIlvl) = EffectiveNumberingOf(p);
                if (pNumId != numId) continue;
                RepointListInstance(p, pIlvl, newNumId.Value);
                repointedUnids.Add((string?)p.Attribute(PtOpenXml.Unid));
            }

            // Flush the body mutation to the part stream immediately — same WASM typed-DOM /
            // XDocument divergence rationale as SetListLevel.
            (ResolvePart(target.PartUri) ?? _doc!.MainDocumentPart!).PutXDocument();
            ClearListNumberingAnnotations();
            InvalidateProjectionCache();
            _ = AnchorIndex();
            var modified = new List<Anchor>();
            foreach (var unid in repointedUnids)
            {
                if (unid is not null && AnchorForUnid(unid, target.PartUri) is { } updated)
                    modified.Add(updated);
            }
            return new EditResult
            {
                Success = true,
                Modified = modified,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// The effective <c>(numId, ilvl)</c> of a paragraph: a direct <c>w:numPr</c> wins per
    /// attribute (its <c>w:numId</c>/<c>w:ilvl</c> each fall back to the pStyle chain via
    /// <see cref="ResolveStyleNumbering"/> when the child is absent). <c>(null, 0)</c> when
    /// neither contributes a numId.
    /// </summary>
    private (int? numId, int ilvl) EffectiveNumberingOf(XElement paragraph)
    {
        var numPr = paragraph.Element(W.pPr)?.Element(W.numPr);
        var directNumId = (int?)numPr?.Element(W.numId)?.Attribute(W.val);
        var directIlvl = (int?)numPr?.Element(W.ilvl)?.Attribute(W.val);
        if (directNumId is not null) return (directNumId, directIlvl ?? 0);
        var (styleNumId, styleIlvl) = ResolveStyleNumbering(paragraph);
        return (styleNumId, directIlvl ?? styleIlvl);
    }

    /// <summary>
    /// Strip the <see cref="ListItemRetriever"/> annotations (<c>ListItemInfo</c> /
    /// <c>LevelNumbers</c> / <c>ContinuationInfo</c>) a previous projection stamped on the live
    /// paragraphs. The retriever re-initializes only paragraphs WITHOUT a <c>ListItemInfo</c>
    /// annotation, so a numbering mutation after a projection would otherwise keep serving the
    /// stale counter vectors — the visible numbers would not restart until save/reopen.
    /// </summary>
    private void ClearListNumberingAnnotations()
    {
        foreach (var part in EnumerateProjectedParts())
        {
            var root = part.GetXDocument().Root;
            if (root is not null) ListItemRetriever.ClearAnnotations(root);
        }
    }

    /// <summary>Point <paramref name="paragraph"/>'s numbering at <paramref name="newNumId"/>,
    /// keeping its effective <paramref name="ilvl"/> — editing the direct <c>w:numPr</c> in place
    /// when one carries a <c>w:numId</c>, else materializing one (the style-derived case).</summary>
    private void RepointListInstance(XElement paragraph, int ilvl, int newNumId)
    {
        var pPr = paragraph.Element(W.pPr);
        var numPr = pPr?.Element(W.numPr);
        if (numPr?.Element(W.numId) is { } numIdEl)
        {
            numIdEl.SetAttributeValue(W.val, newNumId);
            return;
        }
        if (pPr is null) { pPr = new XElement(W.pPr); paragraph.AddFirst(pPr); }
        pPr.Element(W.numPr)?.Remove();
        SetPPrChildInOrder(pPr, new XElement(W.numPr,
            new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
            new XElement(W.numId, new XAttribute(W.val, newNumId))));
    }

    // ─── Tier E: annotations ────────────────────────────────────────────

    /// <summary>
    /// Annotate the range <paramref name="span"/> inside the block addressed by
    /// <paramref name="anchorId"/>. When <paramref name="span"/> is null, the
    /// annotation wraps every inline run of the block. When
    /// <paramref name="annotation"/>.Id is null/empty, a 16-char hex id is
    /// generated. The bookmark name, AnnotatedText, Created, and PageInfoStale
    /// fields of the annotation are always set by this method.
    /// </summary>
    public EditResult AddAnnotation(string anchorId, CharSpan? span, DocumentAnnotation annotation)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (annotation is null)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "annotation is null", anchorId);

        var anchor = FindAnchor(anchorId);
        if (anchor is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Add(_doc!, anchor, span, annotation);
            if (result.Success) InvalidateProjectionCache();
            else _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>Removes an annotation (its bookmark and custom-XML entry) by id.</summary>
    public EditResult RemoveAnnotation(string annotationId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Remove(_doc!, annotationId, CanonicalizeAnchorByUnid);
            if (result.Success) InvalidateProjectionCache();
            else _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>Mutates label/color/author/metadata of an annotation without re-targeting.</summary>
    public EditResult UpdateAnnotation(string annotationId, AnnotationUpdate update)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (update is null)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "update is null");

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Update(_doc!, annotationId, update);
            if (!result.Success) _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>Re-targets an existing annotation to a new anchor + span.</summary>
    public EditResult MoveAnnotation(string annotationId, string newAnchorId, CharSpan? newSpan)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var anchor = FindAnchor(newAnchorId);
        if (anchor is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                $"anchor not found: {newAnchorId}", newAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var result = Internal.AnnotationOps.Move(
                _doc!, annotationId, anchor, newSpan, CanonicalizeAnchorByUnid);
            if (result.Success) InvalidateProjectionCache();
            else _ = _history.PopForUndo();
            return result;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, newAnchorId);
        }
    }

    /// <summary>
    /// Looks up the canonical <see cref="Anchor"/> for a Unid in the current
    /// projection. Used by annotation ops so that the <see cref="EditResult.Modified"/>
    /// anchor matches what <see cref="Project"/>'s AnchorIndex will return on the
    /// next tick — bypasses the local kind/scope classifier in <c>AnnotationOps</c>
    /// drifting from the projector.
    /// </summary>
    private Anchor? CanonicalizeAnchorByUnid(string unid)
    {
        var idx = AnchorIndex();
        return idx.Values.FirstOrDefault(t => t.Unid == unid)?.Anchor;
    }

    // ─── Maintenance / cleanup ───────────────────────────────────────────

    /// <summary>
    /// Remove every <c>w:r</c> in the selected scopes whose only content is a
    /// <c>w:rPr</c> (no text, no tabs, no breaks, no field/footnote/comment
    /// references). Generally useful after any workflow that deletes inline
    /// content — accepting tracked changes, removing footnotes/comments, run-text
    /// refactors — and leaves behind formatting-only runs that the document
    /// model carries but that have no visible effect on rendering.
    /// </summary>
    /// <param name="scopes">Which package parts to compact. Defaults to
    /// <see cref="ProjectionScopes.All"/>.</param>
    /// <returns>How many runs were removed. <c>0</c> means the document was
    /// already compact within the selected scopes.</returns>
    /// <remarks>
    /// One pre-op snapshot is recorded; <see cref="Undo"/> rolls every removal
    /// back together. Block-level anchors (paragraphs / headings / list items /
    /// tables / table cells) are unaffected — runs aren't part of the
    /// <see cref="MarkdownProjection.AnchorIndex"/>.
    /// </remarks>
    public CompactResult CompactRuns(ProjectionScopes scopes = ProjectionScopes.All)
    {
        ThrowIfDisposed();
        _history.RecordPreOp(TakeSnapshot());

        int removed = 0;
        foreach (var part in EnumerateProjectedPartsForScopes(scopes))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            // Materialize before mutating — Remove() during enumeration is unsafe.
            foreach (var r in root.Descendants(W.r).ToList())
            {
                if (IsEmptyRun(r))
                {
                    r.Remove();
                    removed++;
                }
            }
            part.PutXDocument();
        }
        if (removed > 0) InvalidateProjectionCache();
        return new CompactResult { RunsRemoved = removed };
    }

    private static bool IsEmptyRun(XElement r)
    {
        foreach (var child in r.Elements())
        {
            if (child.Name == W.rPr) continue;
            // any other child (w:t, w:tab, w:br, w:footnoteReference, …) is meaningful
            return false;
        }
        return true;
    }

    private IEnumerable<OpenXmlPart> EnumerateProjectedPartsForScopes(ProjectionScopes scopes)
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) yield break;
        if (scopes.HasFlag(ProjectionScopes.Body)) yield return main;
        if (scopes.HasFlag(ProjectionScopes.Headers))
            foreach (var h in main.HeaderParts) yield return h;
        if (scopes.HasFlag(ProjectionScopes.Footers))
            foreach (var f in main.FooterParts) yield return f;
        if (scopes.HasFlag(ProjectionScopes.Footnotes) && main.FootnotesPart is not null)
            yield return main.FootnotesPart;
        if (scopes.HasFlag(ProjectionScopes.Endnotes) && main.EndnotesPart is not null)
            yield return main.EndnotesPart;
        if (scopes.HasFlag(ProjectionScopes.Comments) && main.WordprocessingCommentsPart is not null)
            yield return main.WordprocessingCommentsPart;
    }

    // ─── Undo / Redo ─────────────────────────────────────────────────────

    public bool Undo()
    {
        if (_disposed) return false;
        var (preOp, ok) = _history.PopForUndo();
        if (!ok) return false;
        _history.RecordForRedo(TakeSnapshot());
        RestoreSnapshot(preOp);
        return true;
    }

    public bool Redo()
    {
        if (_disposed) return false;
        var (postOp, ok) = _history.PopForRedo();
        if (!ok) return false;
        _history.PushBackForUndo(TakeSnapshot());
        RestoreSnapshot(postOp);
        return true;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _doc?.Dispose();
        _stream?.Dispose();
        _doc = null;
        _stream = null;
    }

    // ─── Internal mutation helpers (used by tier methods landing in later phases) ───

    internal void InvalidateProjectionCache()
    {
        _cachedProjection = null;
        _cachedAnchorIndex = null;
    }

    /// <summary>
    /// A per-part XML snapshot covering every part the projector / mutation ops walk.
    /// Originally captured only <c>MainDocumentPart</c>, but any cross-part mutation
    /// (footnote definition removal + body reference cleanup, comment range marker
    /// stripping, Save's Unid-strip pass) needs to round-trip all parts — otherwise
    /// undo or the Save restore would leak structural changes into peer parts.
    /// </summary>
    /// <param name="Parts">Per-URI XML content of every snapshot-scoped part (content restore).</param>
    /// <param name="HeaderFooterParts">Relationship id + kind + URI of each header/footer part that
    /// existed at snapshot time. Drives create/delete reconciliation in <see cref="RestoreSnapshot"/>
    /// so ops that add a header/footer part (SetHeaderText/SetFooterText) undo/redo cleanly; the
    /// content is read back from <see cref="Parts"/> by URI when a part must be re-created.</param>
    /// <param name="NoteParts">The same, for the footnotes/endnotes parts, which
    /// InsertFootnote/InsertEndnote create on a document that had no notes.</param>
    /// <param name="CommentParts">The same, for the comments part (0 or 1 entries), which
    /// AddComment creates on a document that had no comments.</param>
    /// <param name="CommentThreadingParts">The same, for commentsExtended/commentsIds, which
    /// AddCommentReply/SetCommentResolved create when upgrading a flat comment.</param>
    internal sealed record DocumentSnapshot(
        System.Collections.Generic.IReadOnlyList<(string PartUri, XDocument Xml)> Parts,
        System.Collections.Generic.IReadOnlyList<(string RelId, bool IsHeader, string PartUri)> HeaderFooterParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, bool IsFootnote, string PartUri)> NoteParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, string PartUri)> CommentParts,
        System.Collections.Generic.IReadOnlyList<(string RelId, bool IsCommentsEx, string PartUri)> CommentThreadingParts);

    internal DocumentSnapshot TakeSnapshot()
    {
        var parts = new System.Collections.Generic.List<(string, XDocument)>();
        foreach (var part in EnumerateProjectedPartsForSnapshot())
            parts.Add((part.Uri.ToString(), new XDocument(part.GetXDocument())));

        var hfParts = new System.Collections.Generic.List<(string, bool, string)>();
        var noteParts = new System.Collections.Generic.List<(string, bool, string)>();
        var commentParts = new System.Collections.Generic.List<(string, string)>();
        var commentThreadingParts = new System.Collections.Generic.List<(string, bool, string)>();
        var main = _doc!.MainDocumentPart;
        if (main is not null)
        {
            foreach (var h in main.HeaderParts) hfParts.Add((main.GetIdOfPart(h), true, h.Uri.ToString()));
            foreach (var f in main.FooterParts) hfParts.Add((main.GetIdOfPart(f), false, f.Uri.ToString()));
            if (main.FootnotesPart is not null)
                noteParts.Add((main.GetIdOfPart(main.FootnotesPart), true, main.FootnotesPart.Uri.ToString()));
            if (main.EndnotesPart is not null)
                noteParts.Add((main.GetIdOfPart(main.EndnotesPart), false, main.EndnotesPart.Uri.ToString()));
            if (main.WordprocessingCommentsPart is not null)
                commentParts.Add((main.GetIdOfPart(main.WordprocessingCommentsPart), main.WordprocessingCommentsPart.Uri.ToString()));
            if (main.WordprocessingCommentsExPart is not null)
                commentThreadingParts.Add((main.GetIdOfPart(main.WordprocessingCommentsExPart), true,
                    main.WordprocessingCommentsExPart.Uri.ToString()));
            if (main.WordprocessingCommentsIdsPart is not null)
                commentThreadingParts.Add((main.GetIdOfPart(main.WordprocessingCommentsIdsPart), false,
                    main.WordprocessingCommentsIdsPart.Uri.ToString()));
        }
        return new DocumentSnapshot(parts, hfParts, noteParts, commentParts, commentThreadingParts);
    }

    internal void RestoreSnapshot(DocumentSnapshot snapshot)
    {
        var byUri = snapshot.Parts.ToDictionary(p => p.PartUri, p => p.Xml);

        // Restore content for all parts that exist in both snapshot and document.
        // Scoped via EnumerateProjectedPartsForSnapshot — only the annotations
        // CustomXmlPart participates here; other CustomXmlParts (SharePoint
        // metadata, SDT data-binding parts, inkml, …) are intentionally outside
        // the snapshot scope.
        foreach (var part in EnumerateProjectedPartsForSnapshot())
        {
            if (!byUri.TryGetValue(part.Uri.ToString(), out var xml)) continue;
            part.PutXDocument(new XDocument(xml));
        }

        var main = _doc!.MainDocumentPart;

        // Header/footer part create/delete reconcile: SetHeaderText/SetFooterText can add a
        // HeaderPart/FooterPart, so undo/redo must delete the parts the snapshot doesn't have and
        // re-create (with the snapshot's relationship id, so the restored sectPr reference resolves)
        // the ones it does. Content restore above already handled parts present in both by URI.
        if (main is not null)
        {
            ReconcileHeaderFooterParts(main, snapshot, byUri);
            // Same reconcile for the footnotes/endnotes parts, which InsertFootnote/InsertEndnote
            // create on a document that had no notes.
            ReconcileNoteParts(main, snapshot, byUri);
            // And for the comments part, which AddComment creates on a document that had no comments.
            ReconcileCommentsPart(main, snapshot, byUri);
            // Reply/resolve can introduce commentsExtended/commentsIds; reconcile their topology
            // after restoring the base comments part.
            ReconcileCommentThreadingParts(main, snapshot, byUri);
        }

        // The annotations CustomXmlPart is reconciled the same way (its own factory) — see
        // EnumerateProjectedPartsForSnapshot for why AddCustomXmlPart(CustomXml) is unsafe for
        // non-annotation custom-xml parts (wrong content type, no CustomXmlPropertiesPart partner).
        if (main is not null)
        {
            var annotationsPart = Internal.AnnotationsCustomXml.Find(_doc);
            var snapshotAnnotationsUri = snapshot.Parts
                .FirstOrDefault(p => p.PartUri.StartsWith("/customXml/", StringComparison.OrdinalIgnoreCase))
                .PartUri;

            // Undo direction: snapshot has no annotations part but the live doc
            // does → forward-op created it, roll it back by deleting.
            if (annotationsPart is not null
                && !byUri.ContainsKey(annotationsPart.Uri.ToString()))
            {
                main.DeletePart(annotationsPart);
                annotationsPart = null;
            }

            // Redo direction: snapshot has an annotations part but the live doc
            // doesn't → undo previously removed it, restore by re-adding.
            if (annotationsPart is null && snapshotAnnotationsUri is not null
                && byUri.TryGetValue(snapshotAnnotationsUri, out var annXml))
            {
                var newPart = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                newPart.PutXDocument(new XDocument(annXml));
            }
        }

        InvalidateProjectionCache();
    }

    /// <summary>
    /// Reconcile the live document's header/footer parts against <paramref name="snapshot"/>:
    /// delete parts created since the snapshot (relationship id present live, absent in snapshot)
    /// and re-create parts removed since it (present in snapshot, absent live) with their original
    /// relationship id + content, so the just-restored sectPr <c>headerReference</c>/<c>footerReference</c>
    /// resolves. Parts present in both keep their content (restored by URI in <see cref="RestoreSnapshot"/>).
    /// </summary>
    private static void ReconcileHeaderFooterParts(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, (bool IsHeader, string PartUri)>(StringComparer.Ordinal);
        foreach (var (relId, isHeader, partUri) in snapshot.HeaderFooterParts)
            snapByRel[relId] = (isHeader, partUri);

        // Live header/footer parts keyed by relationship id (materialized so we can DeletePart
        // without mutating a collection we're iterating).
        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        foreach (var h in main.HeaderParts) live[main.GetIdOfPart(h)] = h;
        foreach (var f in main.FooterParts) live[main.GetIdOfPart(f)] = f;

        // Delete parts the snapshot doesn't know about (undo of a create).
        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        // Re-create parts the snapshot has but the live doc lost (redo of a create / undo of a delete).
        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value.PartUri, out var xml)) continue;
            OpenXmlPart np = kv.Value.IsHeader
                ? main.AddNewPart<HeaderPart>(kv.Key)
                : main.AddNewPart<FooterPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// The <see cref="ReconcileHeaderFooterParts"/> twin for the footnotes/endnotes parts: delete a
    /// part created since <paramref name="snapshot"/> (undo of an InsertFootnote/InsertEndnote that
    /// introduced notes) and re-create one the live document has since lost (redo), keeping the
    /// original relationship id so the package relationship the restored XML expects still resolves.
    /// Parts present in both keep their content, restored by URI in <see cref="RestoreSnapshot"/>.
    /// </summary>
    private static void ReconcileNoteParts(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, (bool IsFootnote, string PartUri)>(StringComparer.Ordinal);
        foreach (var (relId, isFootnote, partUri) in snapshot.NoteParts)
            snapByRel[relId] = (isFootnote, partUri);

        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        if (main.FootnotesPart is not null) live[main.GetIdOfPart(main.FootnotesPart)] = main.FootnotesPart;
        if (main.EndnotesPart is not null) live[main.GetIdOfPart(main.EndnotesPart)] = main.EndnotesPart;

        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value.PartUri, out var xml)) continue;
            OpenXmlPart np = kv.Value.IsFootnote
                ? main.AddNewPart<FootnotesPart>(kv.Key)
                : main.AddNewPart<EndnotesPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// The <see cref="ReconcileNoteParts"/> twin for the comments part: delete a part created
    /// since <paramref name="snapshot"/> (undo of the AddComment that introduced comments) and
    /// re-create one the live document has since lost (redo), keeping the original relationship
    /// id. Content for a part present in both is restored by URI in <see cref="RestoreSnapshot"/>.
    /// </summary>
    private static void ReconcileCommentsPart(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (relId, partUri) in snapshot.CommentParts)
            snapByRel[relId] = partUri;

        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        if (main.WordprocessingCommentsPart is not null)
            live[main.GetIdOfPart(main.WordprocessingCommentsPart)] = main.WordprocessingCommentsPart;

        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value, out var xml)) continue;
            var np = main.AddNewPart<WordprocessingCommentsPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    /// <summary>
    /// Create/delete reconciliation for <c>commentsExtended.xml</c> and
    /// <c>commentsIds.xml</c>. These used to be content-only snapshot parts because no session op
    /// authored them; AddCommentReply/SetCommentResolved can now create either/both, so undo must
    /// remove those parts and redo must restore their original relationship ids and XML.
    /// </summary>
    private static void ReconcileCommentThreadingParts(
        MainDocumentPart main, DocumentSnapshot snapshot,
        System.Collections.Generic.Dictionary<string, XDocument> byUri)
    {
        var snapByRel = new System.Collections.Generic.Dictionary<string, (bool IsCommentsEx, string PartUri)>(StringComparer.Ordinal);
        foreach (var (relId, isCommentsEx, partUri) in snapshot.CommentThreadingParts)
            snapByRel[relId] = (isCommentsEx, partUri);

        var live = new System.Collections.Generic.Dictionary<string, OpenXmlPart>(StringComparer.Ordinal);
        if (main.WordprocessingCommentsExPart is not null)
            live[main.GetIdOfPart(main.WordprocessingCommentsExPart)] = main.WordprocessingCommentsExPart;
        if (main.WordprocessingCommentsIdsPart is not null)
            live[main.GetIdOfPart(main.WordprocessingCommentsIdsPart)] = main.WordprocessingCommentsIdsPart;

        foreach (var kv in live)
            if (!snapByRel.ContainsKey(kv.Key))
                main.DeletePart(kv.Value);

        foreach (var kv in snapByRel)
        {
            if (live.ContainsKey(kv.Key)) continue;
            if (!byUri.TryGetValue(kv.Value.PartUri, out var xml)) continue;
            OpenXmlPart np = kv.Value.IsCommentsEx
                ? main.AddNewPart<WordprocessingCommentsExPart>(kv.Key)
                : main.AddNewPart<WordprocessingCommentsIdsPart>(kv.Key);
            np.PutXDocument(new XDocument(xml));
        }
    }

    internal int NextRevisionId() => System.Threading.Interlocked.Increment(ref _revisionCounter);

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocxSession));
    }

    // ─── Mutation helpers (shared across tiers) ───────────────────────────

    /// <summary>
    /// The per-op patch, or <c>null</c> when <see cref="DocxSessionSettings.EmitMarkdownPatch"/>
    /// is off — every mutation's <c>Patch =</c> site routes through here so the opt-out
    /// cannot be missed by a new op.
    /// </summary>
    private MarkdownPatch? PatchFor(AnchorTarget target) =>
        _settings.EmitMarkdownPatch ? ProjectScope(target) : null;

    internal MarkdownPatch ProjectScope(AnchorTarget target)
    {
        // Phase 3 implementation: re-project the whole document. The patch contract
        // (smallest enclosing block) is honored by ScopeAnchorId; the markdown payload
        // is the full projection until we optimize this in a later phase.
        //
        // Every Patch site runs AFTER the op's InvalidateProjectionCache, so the fresh
        // projection built here IS the post-op state — cache it. Without this, a
        // default-settings caller pays this Convert per op AND a second index build on
        // the next op's FindAnchor.
        var fresh = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
        _cachedProjection = fresh;
        _cachedAnchorIndex = null;
        return new MarkdownPatch(target.Anchor.Id, fresh.Markdown);
    }

    // Zero-width, semantically-significant inline markers that must survive ReplaceText.
    // Discarding them silently destroys bookmark/comment/permission ranges that point
    // into the paragraph from other parts of the document.
    private static readonly HashSet<XName> PreservedMarkerNames = new()
    {
        W.bookmarkStart, W.bookmarkEnd,
        W.commentRangeStart, W.commentRangeEnd, W.commentReference,
        W.permStart, W.permEnd,
        W.proofErr,
    };

    // Inline references that point into another document part (the footnotes/endnotes
    // part). Like comment references, they are zero-width but semantically significant:
    // dropping the body-side <w:footnoteReference w:id="N"/> orphans the note definition
    // and silently loses content on a text edit (issue B3). Unlike the bare-child markers
    // above, these live inside a <w:r>, so they are detected via IsNoteRefOnlyRun.
    private static readonly HashSet<XName> NoteReferenceNames = new()
    {
        W.footnoteReference, W.endnoteReference,
    };

    // True for a run whose only meaningful (non-rPr) content is a footnote/endnote
    // reference — i.e. it carries no visible text. Such a run is a preserved marker;
    // a run that mixes a note ref with text is ordinary content and is replaced.
    private static bool IsNoteRefOnlyRun(XElement e)
    {
        if (e.Name != W.r) return false;
        bool sawNoteRef = false;
        foreach (var child in e.Elements())
        {
            if (child.Name == W.rPr) continue;
            if (NoteReferenceNames.Contains(child.Name)) { sawNoteRef = true; continue; }
            return false; // any other content (w:t, w:tab, w:br, …) ⇒ ordinary run
        }
        return sawNoteRef;
    }

    private static (List<XElement> pre, List<XElement> post) ExtractWrappingMarkers(XElement paragraph)
    {
        var children = paragraph.Elements().Where(e => e.Name != W.pPr).ToList();
        // Position note-ref-only runs relative to the runs that actually carry text, so a
        // leading reference sorts before the replacement and a trailing one after it.
        int firstTextIdx = children.FindIndex(c => IsInlineChild(c) && !IsNoteRefOnlyRun(c));
        int lastTextIdx = children.FindLastIndex(c => IsInlineChild(c) && !IsNoteRefOnlyRun(c));
        var pre = new List<XElement>();
        var post = new List<XElement>();
        for (int i = 0; i < children.Count; i++)
        {
            var c = children[i];
            if (!PreservedMarkerNames.Contains(c.Name) && !IsNoteRefOnlyRun(c)) continue;
            if (firstTextIdx < 0 || i < firstTextIdx) pre.Add(c);
            else if (i > lastTextIdx) post.Add(c);
            else pre.Add(c); // interleaved → wrap from the start (best-effort)
        }
        return (pre, post);
    }

    /// <summary>
    /// If <paramref name="paragraph"/> carries a resolvable <c>w:numPr</c> auto-number
    /// (e.g. <c>"1."</c>, <c>"Fourth"</c>), strip a matching leading prefix from
    /// <paramref name="payload"/> plus one optional separator character (ASCII space,
    /// tab, or NBSP — matching the projector's emission and the common variants an
    /// agent might use). Idempotent when the prefix isn't present.
    /// </summary>
    private string StripResolvedAutoNumberPrefix(XElement paragraph, string payload)
    {
        if (string.IsNullOrEmpty(payload)) return payload;
        // ListItemRetrieverSettings is internal to the projector; pass null so the
        // resolver uses defaults that match what the projector itself emits.
        var prefix = Internal.ListNumberResolver.Resolve(paragraph, _doc!);
        if (string.IsNullOrEmpty(prefix)) return payload;
        if (!payload.StartsWith(prefix, StringComparison.Ordinal)) return payload;

        var after = payload.Substring(prefix.Length);
        if (after.Length > 0 && (after[0] == ' ' || after[0] == '\t' || after[0] == ' '))
            after = after.Substring(1);
        return after;
    }

    private static void ApplyReplaceTextAccept(XElement paragraph, IReadOnlyList<Internal.ParsedBlock> blocks)
    {
        var pPr = paragraph.Element(W.pPr);
        var (preMarkers, postMarkers) = ExtractWrappingMarkers(paragraph);
        paragraph.RemoveNodes();
        if (pPr is not null) paragraph.Add(pPr);
        foreach (var m in preMarkers) paragraph.Add(m);
        if (blocks.Count > 0)
            foreach (var run in blocks[0].RunElements)
                paragraph.Add(new XElement(run));
        foreach (var m in postMarkers) paragraph.Add(m);
    }

    private void ApplyReplaceTextTracked(XElement paragraph, IReadOnlyList<Internal.ParsedBlock> blocks)
    {
        var author = _revisionAuthor ?? "docxodus";
        var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

        // Note references (footnote/endnote) are zero-width, semantically-significant
        // markers that must survive the edit on BOTH accept and reject — they must not
        // be swept into the w:del (issue B3). Pull the note-ref-only runs out (recording
        // whether each sat before or after the visible text) so we can replace them
        // around the del/ins. Bare-child markers (bookmark/comment ranges) are left in
        // place, exactly as before. ExtractWrappingMarkers gives us the leading/trailing
        // split relative to the text runs.
        var (preMarkers, postMarkers) = ExtractWrappingMarkers(paragraph);
        var preNoteRefs = preMarkers.Where(IsNoteRefOnlyRun).ToList();
        var postNoteRefs = postMarkers.Where(IsNoteRefOnlyRun).ToList();
        foreach (var m in preNoteRefs) m.Remove();
        foreach (var m in postNoteRefs) m.Remove();

        // Wrap remaining existing runs (the visible text) in w:del (converting w:t to w:delText).
        var existingRuns = paragraph.Elements(W.r).ToList();
        XElement? del = null;
        if (existingRuns.Count > 0)
        {
            del = new XElement(W.del,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date));
            foreach (var run in existingRuns)
            {
                run.Remove();
                foreach (var t in run.Elements(W.t).ToList())
                {
                    var dt = new XElement(W.delText,
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        (string)t);
                    t.ReplaceWith(dt);
                }
                del.Add(run);
            }
        }

        XElement? ins = null;
        if (blocks.Count > 0 && blocks[0].RunElements.Count > 0)
        {
            ins = new XElement(W.ins,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date));
            foreach (var run in blocks[0].RunElements)
                ins.Add(new XElement(run));
        }

        foreach (var m in preNoteRefs) paragraph.Add(m);
        if (del is not null) paragraph.Add(del);
        if (ins is not null) paragraph.Add(ins);
        foreach (var m in postNoteRefs) paragraph.Add(m);
    }

    private void WrapRunsInDel(XElement element)
    {
        var author = _revisionAuthor ?? "docxodus";
        var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
        foreach (var run in element.Elements(W.r).ToList())
        {
            run.Remove();
            foreach (var t in run.Elements(W.t).ToList())
                t.ReplaceWith(new XElement(W.delText,
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    (string)t));
            var del = new XElement(W.del,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date),
                run);
            element.Add(del);
        }
    }

    /// <summary>
    /// Marks a whole paragraph as a tracked deletion: wraps every direct-child run in
    /// <c>w:del</c> (via <see cref="WrapRunsInDel"/>) AND marks the paragraph mark
    /// itself by adding <c>w:del</c> inside <c>w:pPr/w:rPr</c>. The combination tells
    /// Word the entire paragraph — content plus paragraph break — is a tracked deletion,
    /// so accepting the change actually removes the paragraph (instead of leaving an
    /// empty paragraph behind, which is what <see cref="WrapRunsInDel"/> alone produces).
    /// </summary>
    private void MarkParagraphAsTrackedDeleted(XElement paragraph)
    {
        WrapRunsInDel(paragraph);

        var pPr = paragraph.Element(W.pPr);
        if (pPr is null)
        {
            pPr = new XElement(W.pPr);
            paragraph.AddFirst(pPr);
        }
        var rPr = pPr.Element(W.rPr);
        if (rPr is null)
        {
            rPr = new XElement(W.rPr);
            pPr.AddFirst(rPr);
        }
        if (rPr.Element(W.del) is null)
        {
            var author = _revisionAuthor ?? "docxodus";
            var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
            rPr.Add(new XElement(W.del,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date)));
        }
    }

    /// <summary>
    /// Marks a whole table as a tracked deletion: every row gets a <c>w:trPr/w:del</c>
    /// marker (Word's row-deletion convention — there is no table-level "delete" markup),
    /// and every paragraph inside every cell is treated like
    /// <see cref="MarkParagraphAsTrackedDeleted"/>. Nested tables recurse.
    /// </summary>
    private void MarkTableAsTrackedDeleted(XElement table)
    {
        var author = _revisionAuthor ?? "docxodus";
        var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

        foreach (var row in table.Elements(W.tr))
        {
            var trPr = row.Element(W.trPr);
            if (trPr is null)
            {
                trPr = new XElement(W.trPr);
                row.AddFirst(trPr);
            }
            if (trPr.Element(W.del) is null)
            {
                trPr.Add(new XElement(W.del,
                    new XAttribute(W.id, NextRevisionId()),
                    new XAttribute(W.author, author),
                    new XAttribute(W.date, date)));
            }

            foreach (var cell in row.Elements(W.tc))
            {
                foreach (var child in cell.Elements().ToList())
                {
                    if (child.Name == W.p)
                        MarkParagraphAsTrackedDeleted(child);
                    else if (child.Name == W.tbl)
                        MarkTableAsTrackedDeleted(child);
                }
            }
        }
    }

    private void PromoteHyperlinkRelationships(XElement paragraph)
    {
        var main = _doc!.MainDocumentPart!;
        // Reuse an existing relationship when the same URL has already been registered.
        // Without dedup, every ReplaceText with a link adds a fresh rId; an agent loop
        // that edits the same paragraph N times grows the .rels file unboundedly.
        var existing = main.HyperlinkRelationships
            .GroupBy(rl => rl.Uri.ToString())
            .ToDictionary(g => g.Key, g => g.First().Id);
        foreach (var link in paragraph.Descendants(W.hyperlink).ToList())
        {
            var hrefAttr = link.Attribute(Internal.MarkdownPayloadParser.HrefAttr);
            if (hrefAttr is null) continue;
            var url = hrefAttr.Value;
            string relId;
            if (existing.TryGetValue(url, out var foundId)) relId = foundId;
            else
            {
                var rel = main.AddHyperlinkRelationship(
                    new Uri(url, UriKind.RelativeOrAbsolute), true);
                relId = rel.Id;
                existing[url] = relId;
            }
            link.SetAttributeValue(R.id, relId);
            hrefAttr.Remove();
        }
    }

    private static void ApplyFormatToRun(XElement run, FormatOp op)
    {
        var rPr = run.Element(W.rPr);
        if (rPr is null) { rPr = new XElement(W.rPr); run.AddFirst(rPr); }

        static void Toggle(XElement rPr, XName name, bool? set)
        {
            if (set is null) return;
            var existing = rPr.Element(name);
            if (set.Value)
            {
                // Turn the property ON. A run may already carry an explicit OFF element
                // (e.g. Google Docs stamps <w:b w:val="0"/> on every run); just adding a new
                // element when one is "missing" would leave that w:val="0" in place and the
                // toggle would silently do nothing. Normalize: drop the w:val so the bare
                // element (<w:b/>) means on; add one only when truly absent.
                if (existing is null) rPr.Add(new XElement(name));
                else existing.Attribute(W.val)?.Remove();
            }
            else existing?.Remove();
        }

        Toggle(rPr, W.b, op.Bold);
        Toggle(rPr, W.i, op.Italic);
        Toggle(rPr, W.strike, op.Strike);

        if (op.Underline is true)
        {
            rPr.Element(W.u)?.Remove();
            rPr.Add(new XElement(W.u, new XAttribute(W.val, "single")));
        }
        else if (op.Underline is false) rPr.Element(W.u)?.Remove();

        if (op.Code is true)
        {
            rPr.Element(W.rStyle)?.Remove();
            rPr.Add(new XElement(W.rStyle, new XAttribute(W.val, "Code")));
        }
        else if (op.Code is false) rPr.Element(W.rStyle)?.Remove();

        if (op.Color is not null)
        {
            rPr.Element(W.color)?.Remove();
            if (op.Color.Length > 0)
                rPr.Add(new XElement(W.color, new XAttribute(W.val, op.Color)));
        }

        if (op.RunStyle is not null)
        {
            rPr.Element(W.rStyle)?.Remove();
            if (op.RunStyle.Length > 0)
                rPr.Add(new XElement(W.rStyle, new XAttribute(W.val, op.RunStyle)));
        }

        if (op.VertAlign is not null)
        {
            rPr.Element(W.vertAlign)?.Remove();
            var v = op.VertAlign switch
            {
                "super" => "superscript",
                "sub" => "subscript",
                "none" or "baseline" => "",
                _ => op.VertAlign,
            };
            if (v.Length > 0)
            {
                if (v is not ("superscript" or "subscript"))
                    throw new ArgumentException($"invalid vertAlign: {op.VertAlign}");
                rPr.Add(new XElement(W.vertAlign, new XAttribute(W.val, v)));
            }
        }

        if (op.FontSizePts is { } pts)
        {
            // w:sz / w:szCs are half-points. Clearing (<= 0) drops the explicit size so the run
            // inherits the style/default size again.
            rPr.Element(W.sz)?.Remove();
            rPr.Element(W.szCs)?.Remove();
            if (pts > 0)
            {
                var halfPts = ((int)System.Math.Round(pts * 2, System.MidpointRounding.AwayFromZero))
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);
                rPr.Add(new XElement(W.sz, new XAttribute(W.val, halfPts)));
                rPr.Add(new XElement(W.szCs, new XAttribute(W.val, halfPts)));
            }
        }

        if (op.FontFamily is not null)
        {
            // w:rFonts is the first EG_RPrBase child after an optional w:rStyle, so it must be
            // placed there (a bare rPr.Add would append after w:sz/w:vertAlign → out of schema
            // order). "" clears the explicit font so the run inherits the style/default.
            rPr.Element(W.rFonts)?.Remove();
            if (op.FontFamily.Length > 0)
            {
                var rFonts = new XElement(W.rFonts,
                    new XAttribute(W.ascii, op.FontFamily),
                    new XAttribute(W.hAnsi, op.FontFamily),
                    new XAttribute(W.cs, op.FontFamily));
                var rStyle = rPr.Element(W.rStyle);
                if (rStyle is not null) rStyle.AddAfterSelf(rFonts);
                else rPr.AddFirst(rFonts);
            }
        }
    }

    /// <summary>
    /// Apply a run-format mutation using Word's native tracked-format representation:
    /// the run keeps its new properties while <c>w:rPr/w:rPrChange/w:rPr</c> stores
    /// the old properties for reject. A run may carry only one direct rPrChange. When
    /// it already has one, preserve that marker (including its original baseline and
    /// attribution) and fold the new formatting into the same pending revision; replacing
    /// its baseline with the intermediate format would make reject-all stop halfway.
    /// </summary>
    private void ApplyFormatToRunTracked(
        XElement run, FormatOp op, string revisionAuthor, string revisionDate)
    {
        var originalRPr = run.Element(W.rPr);
        var originalRPrClone = originalRPr is null ? null : new XElement(originalRPr);
        var oldProperties = SnapshotRunPropertiesForRevision(originalRPr);

        // rPrChange is the final CT_RPr child. Detach it while ApplyFormatToRun edits
        // properties (several setters append) and re-append it below. Taking only the
        // first also prevents malformed duplicate markers from becoming nested/stacked.
        var existingChanges = originalRPr?.Elements(W.rPrChange).ToList()
            ?? new List<XElement>();
        foreach (var change in existingChanges) change.Remove();
        var existingChange = existingChanges.FirstOrDefault();
        var archivedProperties = existingChange is null
            ? null
            : SnapshotRunPropertiesForRevision(existingChange.Element(W.rPr));

        try
        {
            ApplyFormatToRun(run, op);
        }
        catch
        {
            RestoreRunProperties(run, originalRPrClone);
            throw;
        }

        var currentRPr = run.Element(W.rPr)!;
        var newProperties = SnapshotRunPropertiesForRevision(currentRPr);
        bool changed = !RunPropertiesEquivalentForRevision(oldProperties, newProperties);

        if (archivedProperties is not null
            && RunPropertiesEquivalentForRevision(archivedProperties, newProperties))
        {
            // Editing a pending format change back to its archived baseline resolves
            // that portion of the change. Keeping rPrChange here would leave a phantom
            // revision whose accept and reject results are identical. Reuse the stored
            // baseline XML so lexical-only differences do not survive as document churn.
            RestoreRunProperties(run,
                archivedProperties.HasElements
                    || archivedProperties.Attributes().Any(a => !a.IsNamespaceDeclaration)
                    ? archivedProperties
                    : null);
            return;
        }

        if (!changed)
        {
            // A no-op must not manufacture a format revision OR normalize/reorder the
            // caller's existing XML as a side effect. Put the exact rPr back.
            RestoreRunProperties(run, originalRPrClone);
            return;
        }

        // ApplyFormatToRun historically appends several properties. Once rPrChange makes
        // schema validity externally observable, normalize the changed outer rPr to the
        // canonical CT_RPr order before placing the revision marker last.
        var orderedRPr = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(currentRPr);
        currentRPr.ReplaceNodes(orderedRPr.Nodes());

        if (existingChange is not null)
        {
            currentRPr.Add(existingChange);
            return;
        }

        currentRPr.Add(new XElement(W.rPrChange,
            new XAttribute(W.id, NextRevisionId()),
            new XAttribute(W.author, revisionAuthor),
            new XAttribute(W.date, revisionDate),
            oldProperties));
    }

    /// <summary>
    /// Clone the direct run properties suitable for the inner payload of rPrChange.
    /// Existing change markup is deliberately excluded (CT_RPrOriginal cannot contain
    /// another rPrChange), as is projector-only Unid bookkeeping.
    /// </summary>
    private static XElement SnapshotRunPropertiesForRevision(XElement? rPr)
    {
        if (rPr is null) return new XElement(W.rPr);

        var snapshot = new XElement(rPr);
        snapshot.Descendants(W.rPrChange).Remove();
        foreach (var element in snapshot.DescendantsAndSelf())
            element.Attributes()
                .Where(a => a.Name.Namespace == PtOpenXml.pt)
                .Remove();
        return snapshot;
    }

    /// <summary>Compare run properties by their schema order and normalize the lexical
    /// true spellings of the on/off toggles ApplyFormat writes as bare elements, bare
    /// underline as <c>single</c>, and canonical lexical forms for color/half-point
    /// values. This keeps semantically identical writes (for example
    /// <c>w:b w:val="true"</c> → <c>w:b</c>, <c>w:u</c> →
    /// <c>w:u w:val="single"</c>, or remove/re-add of an unchanged color) out of the
    /// review pane.</summary>
    private static bool RunPropertiesEquivalentForRevision(XElement left, XElement right)
    {
        static XElement Normalize(XElement source)
        {
            var ordered = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(source);
            foreach (var name in new[] { W.b, W.i, W.strike })
            {
                foreach (var property in ordered.Elements(name).ToList())
                {
                    var value = (string?)property.Attribute(W.val);
                    if (value is null || value is "1"
                        || value.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("on", StringComparison.OrdinalIgnoreCase))
                    {
                        property.Attribute(W.val)?.Remove();
                    }
                    else if (value is "0"
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("off", StringComparison.OrdinalIgnoreCase))
                    {
                        property.Remove();
                    }
                }
            }

            foreach (var underline in ordered.Elements(W.u).ToList())
            {
                var value = (string?)underline.Attribute(W.val);
                if (string.IsNullOrEmpty(value)
                    || value.Equals("single", StringComparison.OrdinalIgnoreCase))
                {
                    underline.SetAttributeValue(W.val, "single");
                }
                else if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    underline.Remove();
                }
            }

            foreach (var vertAlign in ordered.Elements(W.vertAlign).ToList())
            {
                var value = (string?)vertAlign.Attribute(W.val);
                if (value is not null
                    && value.Equals("baseline", StringComparison.OrdinalIgnoreCase))
                {
                    vertAlign.Remove();
                }
            }

            foreach (var color in ordered.Elements(W.color))
            {
                var value = (string?)color.Attribute(W.val);
                if (value is null) continue;
                if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
                    color.SetAttributeValue(W.val, "auto");
                else if (value.Length == 6 && value.All(Uri.IsHexDigit))
                    color.SetAttributeValue(W.val, value.ToUpperInvariant());
            }

            foreach (var name in new[] { W.sz, W.szCs })
            {
                foreach (var size in ordered.Elements(name))
                {
                    var value = (string?)size.Attribute(W.val);
                    if (uint.TryParse(value, System.Globalization.NumberStyles.None,
                            System.Globalization.CultureInfo.InvariantCulture, out var parsed))
                    {
                        size.SetAttributeValue(W.val, parsed.ToString(
                            System.Globalization.CultureInfo.InvariantCulture));
                    }
                }
            }
            return ordered;
        }

        return XNode.DeepEquals(Normalize(left), Normalize(right));
    }

    private static void RestoreRunProperties(XElement run, XElement? snapshot)
    {
        run.Element(W.rPr)?.Remove();
        if (snapshot is not null) run.AddFirst(snapshot);
    }

    internal static XElement BuildParagraphFromParsedBlock(Internal.ParsedBlock block)
    {
        var p = new XElement(W.p);
        var pPr = new XElement(W.pPr);

        switch (block.Kind)
        {
            case Internal.ParserBlockKind.Heading1:
            case Internal.ParserBlockKind.Heading2:
            case Internal.ParserBlockKind.Heading3:
            case Internal.ParserBlockKind.Heading4:
            case Internal.ParserBlockKind.Heading5:
            case Internal.ParserBlockKind.Heading6:
                {
                    int level = (int)block.Kind - (int)Internal.ParserBlockKind.Heading1 + 1;
                    pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, $"Heading{level}")));
                    break;
                }
            case Internal.ParserBlockKind.Quote:
                pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, "Quote")));
                break;
            case Internal.ParserBlockKind.Code:
                pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, "Code")));
                break;
            // List items: numPr inheritance not auto-injected in v1 — caller can use
            // SetListLevel afterwards if needed. The bare paragraph will project as a
            // normal paragraph until numbering is added.
        }

        if (pPr.HasElements) p.Add(pPr);
        foreach (var run in block.RunElements)
            p.Add(new XElement(run));
        return p;
    }

    internal static string ParserBlockKindToAnchorKind(Internal.ParserBlockKind kind) => kind switch
    {
        Internal.ParserBlockKind.Heading1
            or Internal.ParserBlockKind.Heading2
            or Internal.ParserBlockKind.Heading3
            or Internal.ParserBlockKind.Heading4
            or Internal.ParserBlockKind.Heading5
            or Internal.ParserBlockKind.Heading6 => "h",
        Internal.ParserBlockKind.BulletItem
            or Internal.ParserBlockKind.OrderedItem => "li",
        _ => "p",
    };

    /// <summary>
    /// Mirror the classifier used by <see cref="WmlToMarkdownConverter"/> so the kind
    /// reported in <see cref="EditResult.Created"/> matches what the projector will
    /// emit on the next <see cref="DocxSession.Project"/>. If we used the parser's
    /// kind blindly, a bullet-payload paragraph without a <c>w:numPr</c> would be
    /// reported as "li" but appear as "p" in the projection — a stale anchor id.
    /// </summary>
    internal static string ClassifyParagraphKind(XElement paragraph)
    {
        var pPr = paragraph.Element(W.pPr);
        var styleId = (string?)pPr?.Element(W.pStyle)?.Attribute(W.val);
        if (!string.IsNullOrEmpty(styleId)
            && (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)))
            return "h";
        if (pPr?.Element(W.numPr) is not null) return "li";
        return "p";
    }

    /// <summary>
    /// Classify any block-level XElement to the kind used in anchor ids. Mirrors
    /// the kinds the projector emits — paragraphs go through
    /// <see cref="ClassifyParagraphKind"/>; tables/rows/cells map to their fixed kinds.
    /// Falls back to "p" for unknown shapes.
    /// </summary>
    internal static string ClassifyBlockKind(XElement element)
    {
        if (element.Name == W.p) return ClassifyParagraphKind(element);
        if (element.Name == W.tbl) return "tbl";
        if (element.Name == W.tr) return "tr";
        if (element.Name == W.tc) return "tc";
        return "p";
    }

    /// <summary>
    /// Copy <c>w:numPr</c> from a nearby sibling list item into the new paragraph so
    /// a bullet/ordered-item payload actually renders as part of an existing list.
    /// Walks previous siblings first (closest match first), then next siblings.
    /// No-op when no sibling carries numbering — caller then reports kind="p" via
    /// <see cref="ClassifyParagraphKind"/>.
    /// </summary>
    private static void TryInheritNumPrFromSibling(XElement newParagraph, XElement anchorElement)
    {
        XElement? donorNumPr = null;
        XElement? donorPStyle = null;
        foreach (var sib in anchorElement.ElementsBeforeSelf().Reverse()
                                .Concat(new[] { anchorElement })
                                .Concat(anchorElement.ElementsAfterSelf()))
        {
            if (sib.Name != W.p) continue;
            var nump = sib.Element(W.pPr)?.Element(W.numPr);
            if (nump is null) continue;
            donorNumPr = nump;
            donorPStyle = sib.Element(W.pPr)?.Element(W.pStyle);
            break;
        }
        if (donorNumPr is null) return;

        var pPr = newParagraph.Element(W.pPr);
        if (pPr is null) { pPr = new XElement(W.pPr); newParagraph.AddFirst(pPr); }
        if (pPr.Element(W.numPr) is null) pPr.Add(new XElement(donorNumPr));
        if (donorPStyle is not null && pPr.Element(W.pStyle) is null)
            pPr.AddFirst(new XElement(donorPStyle));
    }

    // Top-level inline children of <w:p> that participate in text flow.
    // Hyperlinks, sdts, fldSimple and smartTag are transparent containers — their
    // descendant runs contribute to the paragraph's visible text. Bookmark/comment
    // markers (zero-width) are tracked separately and not enumerated here.
    private static readonly HashSet<XName> InlineContainerNames = new()
    {
        W.hyperlink, W.sdt, W.fldSimple, W.smartTag,
    };

    private static bool IsInlineChild(XElement e) =>
        e.Name == W.r || InlineContainerNames.Contains(e.Name);

    /// <summary>
    /// All <c>&lt;w:r&gt;</c> elements that contribute to the paragraph's visible text,
    /// in document order — including runs nested inside hyperlinks, sdts, fldSimple,
    /// smartTags. Iterating only <c>Elements(W.r)</c> silently skips hyperlink-internal
    /// runs, which produced the bugs documented in DS080-DS090.
    /// </summary>
    internal static IEnumerable<XElement> InlineRuns(XElement paragraph)
    {
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == W.r) yield return child;
            else if (InlineContainerNames.Contains(child.Name))
                foreach (var run in child.Descendants(W.r))
                    yield return run;
        }
    }

    internal static string ParagraphText(XElement paragraph) =>
        string.Concat(InlineRuns(paragraph).Select(RunText));

    internal static string RunText(XElement run) =>
        string.Concat(run.Elements(W.t).Select(t => (string)t));

    private static int InlineChildTextLength(XElement child) =>
        string.Concat(child.DescendantsAndSelf(W.t).Select(t => (string)t)).Length;

    /// <summary>
    /// If a run straddles <paramref name="offset"/>, split it into two adjacent runs
    /// at that offset. Walks runs inside hyperlinks/sdts/etc. too, so the boundary
    /// is clean regardless of which container the run lives in. The new sibling run
    /// is inserted into the same parent as the original (preserving hyperlink/sdt
    /// membership for the keep-half).
    /// </summary>
    internal static void SplitRunsAtOffset(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var run in InlineRuns(paragraph).ToList())
        {
            var runText = RunText(run);
            if (consumed == offset) return;
            if (consumed + runText.Length <= offset) { consumed += runText.Length; continue; }
            int splitAt = offset - consumed;
            if (splitAt <= 0) return;

            var keep = runText.Substring(0, splitAt);
            var move = runText.Substring(splitAt);

            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"), keep));

            var rPr = run.Element(W.rPr);
            var newRun = new XElement(W.r);
            if (rPr is not null) newRun.Add(new XElement(rPr));
            newRun.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"), move));
            run.AddAfterSelf(newRun);
            return;
        }
    }

    /// <summary>
    /// Ensures no top-level inline child straddles <paramref name="offset"/>: if a
    /// hyperlink (or other splittable container) crosses the boundary, it's split
    /// into two sibling containers sharing the same attributes (e.g. <c>r:id</c>),
    /// each holding half the runs. After this call, <see cref="MoveInlineChildrenAfter"/>
    /// can move whole-child elements without slicing through anything.
    /// </summary>
    internal static void SplitInlineContainersAtOffset(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var child in paragraph.Elements().Where(IsInlineChild).ToList())
        {
            int len = InlineChildTextLength(child);
            if (consumed + len <= offset) { consumed += len; continue; }
            if (consumed == offset) return; // boundary already clean
            int local = offset - consumed;

            if (child.Name == W.hyperlink)
                SplitHyperlinkAt(child, local);
            // For <w:r>: SplitRunsAtOffset already handled it. For sdt/fldSimple/smartTag:
            // treat as atomic — splitting these requires semantic care; the whole element
            // stays with whichever side its leading run lands on.
            return;
        }
    }

    private static void SplitHyperlinkAt(XElement hyperlink, int localOffset)
    {
        // Split runs inside the hyperlink at the local offset (works because SplitRunsAtOffset
        // walks descendants through container types).
        SplitRunsAtOffset(hyperlink, localOffset);

        int consumed = 0;
        var movedRuns = new List<XElement>();
        foreach (var run in hyperlink.Elements(W.r).ToList())
        {
            int len = RunText(run).Length;
            if (consumed >= localOffset) movedRuns.Add(run);
            consumed += len;
        }
        if (movedRuns.Count == 0) return;

        var newLink = new XElement(W.hyperlink);
        foreach (var a in hyperlink.Attributes()) newLink.SetAttributeValue(a.Name, a.Value);
        foreach (var run in movedRuns) { run.Remove(); newLink.Add(run); }
        hyperlink.AddAfterSelf(newLink);
    }

    /// <summary>
    /// Move every paragraph child (inline run/container OR zero-width marker)
    /// whose position is at or past <paramref name="offset"/> from
    /// <paramref name="paragraph"/> into <paramref name="destination"/>. Inline
    /// children advance the position counter by their text length; markers
    /// (bookmarkStart/End, comment range markers, etc.) advance it by 0 and so
    /// inherit the position they're sandwiched between.
    /// </summary>
    internal static void MoveInlineChildrenAfter(XElement paragraph, int offset, XElement destination)
    {
        int consumed = 0;
        var toMove = new List<XElement>();
        foreach (var child in paragraph.Elements().ToList())
        {
            if (child.Name == W.pPr) continue;
            int len = IsInlineChild(child) ? InlineChildTextLength(child) : 0;
            if (consumed >= offset) toMove.Add(child);
            consumed += len;
        }
        foreach (var c in toMove) { c.Remove(); destination.Add(c); }
    }
}
