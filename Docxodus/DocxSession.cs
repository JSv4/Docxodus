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

/// <summary>List membership for <see cref="DocxSession.ApplyListFormat"/>.</summary>
public enum ListFormat { None, Bullet, Decimal }

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

    MalformedXml,
    DisallowedNamespace,
    IncompatibleElementType,
    ValidationFailed,

    NothingToUndo,
    NothingToRedo,

    DuplicateAnnotationId,
    AnnotationNotFound,
    EmptyAnnotationSpan,

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
    private RawDocxOps? _raw;

    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        _settings = settings ?? new DocxSessionSettings();
        _history = new Internal.UndoRing<DocumentSnapshot>(_settings.UndoDepth);
        _stream = new MemoryStream();
        _stream.Write(docxBytes, 0, docxBytes.Length);
        _stream.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);

        if (_settings.CaptureInitialProjection)
            _initialProjection = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
    }

    public Exception? LastInternalError { get; private set; }

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
    internal AnchorTarget? FindAnchor(string? anchorId)
    {
        if (anchorId is null) return null;
        var index = Project().AnchorIndex;
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

    public bool Exists(string anchorId)
    {
        ThrowIfDisposed();
        return FindAnchor(anchorId) is not null;
    }

    public AnchorInfo? GetAnchorInfo(string anchorId)
    {
        ThrowIfDisposed();
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
                Patch = ProjectScope(target),
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
                Patch = ProjectScope(target),
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

    public byte[] Save()
    {
        ThrowIfDisposed();

        if (_settings.PersistAnchorIds)
        {
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return _stream.ToArray();
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
            return _stream.ToArray();
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
            if (_settings.TrackedChanges == TrackedChangeMode.RenderInline)
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
                Patch = ProjectScope(target),
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

        // Word reserves a couple of footnote/endnote definitions (the "separator" and
        // "continuationSeparator" types) for page-rendering scaffolding; they have no
        // user-content meaning and removing them corrupts the doc. Refuse explicitly.
        if (target.Anchor.Kind is "fn" or "en")
        {
            var typeAttr = (string?)element.Attribute(W.type);
            if (typeAttr is "separator" or "continuationSeparator")
                return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                    $"cannot delete a Word-reserved {target.Anchor.Kind} of type='{typeAttr}'", anchorId);
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Tracked-change mode wraps removed runs in w:del — only meaningful for
            // body-level paragraph kinds. fn/en/cmt are structural definitions in
            // their own parts; "tracking" a definition deletion has no Word semantics,
            // so for those we always perform the structural delete.
            if (_settings.TrackedChanges == TrackedChangeMode.RenderInline
                && target.Anchor.Kind is "p" or "h" or "li")
            {
                WrapRunsInDel(element);
                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = new[] { target.Anchor },
                    Patch = ProjectScope(target),
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
            }

            // Collect descendant anchors before removal so the caller knows what's gone.
            var index = Project().AnchorIndex;
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
                Patch = ProjectScope(target),
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
            var index = Project().AnchorIndex;
            bool trackedChanges = _settings.TrackedChanges == TrackedChangeMode.RenderInline;

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
                    Patch = ProjectScope(anchorForPatchScope),
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
                Patch = ProjectScope(anchorForPatchScope),
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
                Patch = ProjectScope(target),
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
                Project().AnchorIndex.Values.FirstOrDefault(t => t.Unid == secondUnid)?.Anchor
                ?? new Anchor(
                    $"{target.Anchor.Kind}:{target.Anchor.Scope}:{secondUnid}",
                    target.Anchor.Kind, target.Anchor.Scope, secondUnid);

            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Created = new[] { secondAnchor },
                Patch = ProjectScope(target),
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
                Patch = ProjectScope(firstTarget),
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
            var freshIndex = Project().AnchorIndex;
            var created = new List<Anchor>();
            foreach (var unid in CollectUnids(parsedXml))
            {
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == unid);
                if (hit is not null) created.Add(hit.Anchor);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = ProjectScope(target),
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
            var freshIndex = Project().AnchorIndex;
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
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid);
                if (hit is not null) modified.Add(hit.Anchor);
            }
            else
            {
                removed.Add(target.Anchor);
            }
            foreach (var unid in newUnids)
            {
                if (unid == target.Unid) continue;
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == unid);
                if (hit is not null) created.Add(hit.Anchor);
            }

            return new EditResult
            {
                Success = true,
                Removed = removed,
                Created = created,
                Modified = modified,
                Patch = ProjectScope(target),
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
                Patch = ProjectScope(target),
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

            int consumed = 0;
            foreach (var run in InlineRuns(element).ToList())
            {
                var runText = RunText(run);
                int runStart = consumed;
                int runEnd = consumed + runText.Length;
                consumed = runEnd;
                if (runEnd <= actualSpan.Start || runStart >= actualSpan.Start + actualSpan.Length) continue;
                ApplyFormatToRun(run, op);
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
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
            var freshIndex = Project().AnchorIndex;
            var updated = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = ProjectScope(target),
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
    /// Set paragraph-level formatting (alignment, indent delta, page-break-before) on the
    /// paragraph the anchor names. Only the non-null fields of <paramref name="op"/> change.
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

            if (op.ClearBorders is true || op.TopBorder is not null || op.BottomBorder is not null)
                ApplyParagraphBorders(pPr, op.TopBorder, op.BottomBorder, op.ClearBorders is true);

            InvalidateProjectionCache();
            var freshIndex = Project().AnchorIndex;
            var updated = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = ProjectScope(target),
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
            var created = Project().AnchorIndex.Values.FirstOrDefault(t => t.Unid == unid)?.Anchor
                ?? new Anchor($"p:{target.Anchor.Scope}:{unid}", "p", target.Anchor.Scope, unid);

            return new EditResult
            {
                Success = true,
                Created = new[] { created },
                Patch = ProjectScope(target),
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
            var index = Project().AnchorIndex;
            var created = new List<Anchor>();
            foreach (var p in cellParagraphs)
            {
                var unid = (string)p.Attribute(PtOpenXml.Unid)!;
                if (index.Values.FirstOrDefault(t => t.Unid == unid)?.Anchor is { } a)
                    created.Add(a);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = ProjectScope(target),
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
        var index = Project().AnchorIndex;
        var result = new List<Anchor>();
        foreach (var para in paras)
        {
            var unid = (string?)para.Attribute(PtOpenXml.Unid);
            if (unid is not null && index.Values.FirstOrDefault(t => t.Unid == unid)?.Anchor is { } a)
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
                Patch = ProjectScope(target!),
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
                Patch = ProjectScope(target!),
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
            var index = Project().AnchorIndex;
            var removed = CellParagraphAnchorsIn(tr!, index);
            if (tbl!.Elements(W.tr).Count() <= 1) { foreach (var a in CellParagraphAnchorsIn(tbl, index)) if (!removed.Contains(a)) removed.Add(a); tbl.Remove(); }
            else tr!.Remove();

            InvalidateProjectionCache();
            return new EditResult { Success = true, Removed = removed, Patch = ProjectScope(target!) };
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
            var index = Project().AnchorIndex;
            var grid = tbl!.Element(W.tblGrid);
            int colCount = grid?.Elements(W.gridCol).Count() ?? tbl.Elements(W.tr).First().Elements(W.tc).Count();

            var removed = new List<Anchor>();
            if (colCount <= 1) { foreach (var a in CellParagraphAnchorsIn(tbl, index)) removed.Add(a); tbl.Remove(); }
            else
            {
                foreach (var tr in tbl.Elements(W.tr).ToList())
                {
                    var cells = tr.Elements(W.tc).ToList();
                    if (colIndex >= cells.Count) continue;
                    foreach (var a in CellParagraphAnchorsIn(cells[colIndex], index)) removed.Add(a);
                    cells[colIndex].Remove();
                }
                var cols = grid?.Elements(W.gridCol).ToList();
                if (cols is not null && colIndex < cols.Count) cols[colIndex].Remove();
            }

            InvalidateProjectionCache();
            return new EditResult { Success = true, Removed = removed, Patch = ProjectScope(target!) };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    /// <summary>The cell-paragraph anchors under <paramref name="scope"/> (a tc/tr/tbl), in document order.</summary>
    private static List<Anchor> CellParagraphAnchorsIn(XElement scope, IReadOnlyDictionary<string, AnchorTarget> index)
    {
        var result = new List<Anchor>();
        foreach (var para in scope.Descendants(W.p))
        {
            var unid = (string?)para.Attribute(PtOpenXml.Unid);
            if (unid is not null && index.Values.FirstOrDefault(t => t.Unid == unid)?.Anchor is { } a)
                result.Add(a);
        }
        return result;
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
            Patch = ProjectScope(target),
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
        var fresh = Project().AnchorIndex;
        var updated = fresh.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;
        return new EditResult
        {
            Success = true,
            Modified = new[] { updated },
            Patch = ProjectScope(target),
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
                var fmt = kind == ListFormat.Bullet ? NumberFormat.Bullet : NumberFormat.Decimal;
                int numId = Internal.NumberingFactory.EnsureNumbering(_doc!, fmt);
                int ilvl = (int?)pPr.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
                pPr.Element(W.numPr)?.Remove();
                SetPPrChildInOrder(pPr, new XElement(W.numPr,
                    new XElement(W.ilvl, new XAttribute(W.val, ilvl)),
                    new XElement(W.numId, new XAttribute(W.val, numId))));
            }

            InvalidateProjectionCache();
            var freshIndex = Project().AnchorIndex;
            var updated = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;
            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
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
        var idx = Project().AnchorIndex;
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

    internal void InvalidateProjectionCache() => _cachedProjection = null;

    /// <summary>
    /// A per-part XML snapshot covering every part the projector / mutation ops walk.
    /// Originally captured only <c>MainDocumentPart</c>, but any cross-part mutation
    /// (footnote definition removal + body reference cleanup, comment range marker
    /// stripping, Save's Unid-strip pass) needs to round-trip all parts — otherwise
    /// undo or the Save restore would leak structural changes into peer parts.
    /// </summary>
    internal sealed record DocumentSnapshot(System.Collections.Generic.IReadOnlyList<(string PartUri, XDocument Xml)> Parts);

    internal DocumentSnapshot TakeSnapshot()
    {
        var parts = new System.Collections.Generic.List<(string, XDocument)>();
        foreach (var part in EnumerateProjectedPartsForSnapshot())
            parts.Add((part.Uri.ToString(), new XDocument(part.GetXDocument())));
        return new DocumentSnapshot(parts);
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

        // TODO: Asymmetric scope, intentional. The undo-time create/delete logic
        // below only handles the annotations CustomXmlPart — see
        // EnumerateProjectedPartsForSnapshot for why AddCustomXmlPart(CustomXml)
        // is unsafe for non-annotation custom-xml parts (wrong content type, no
        // CustomXmlPropertiesPart partner). The same hazard would exist for
        // HeaderParts / FooterParts / FootnotesPart / etc. — but no session op
        // creates or deletes those today, so they're not yet snapshot-scoped here.
        // If a future op starts adding/removing those parts, expand this block
        // (and the snapshot enumeration) to cover them with the correct factory.
        var main = _doc!.MainDocumentPart;
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

    internal int NextRevisionId() => System.Threading.Interlocked.Increment(ref _revisionCounter);

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocxSession));
    }

    // ─── Mutation helpers (shared across tiers) ───────────────────────────

    internal MarkdownPatch ProjectScope(AnchorTarget target)
    {
        // Phase 3 implementation: re-project the whole document. The patch contract
        // (smallest enclosing block) is honored by ScopeAnchorId; the markdown payload
        // is the full projection until we optimize this in a later phase.
        var fresh = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
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
        var author = _settings.RevisionAuthor ?? "docxodus";
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
        var author = _settings.RevisionAuthor ?? "docxodus";
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
            var author = _settings.RevisionAuthor ?? "docxodus";
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
        var author = _settings.RevisionAuthor ?? "docxodus";
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
