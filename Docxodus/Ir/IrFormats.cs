#nullable enable

namespace Docxodus.Ir;

/// <summary>Underline style (`w:u/@w:val`). <see cref="Other"/> covers values the IR does not model individually.</summary>
internal enum IrUnderlineKind { Single, Double, Thick, Dotted, Dashed, Wave, Words, None, Other }

/// <summary>An underline: its <paramref name="Kind"/> and optional color (`w:u/@w:color`, as written).</summary>
internal sealed record IrUnderline(IrUnderlineKind Kind, string? ColorHex);

/// <summary>Vertical run alignment (`w:vertAlign`): subscript or superscript.</summary>
internal enum IrVertAlign { Subscript, Superscript }

/// <summary>Paragraph justification (`w:jc`). <see cref="Other"/> covers values the IR does not model individually.</summary>
internal enum IrJustification { Left, Center, Right, Both, Distribute, Other }

/// <summary>Line-spacing rule (`w:spacing/@w:lineRule`).</summary>
internal enum IrLineSpacingRule { Auto, AtLeast, Exact }

/// <summary>Line spacing: the value in twips plus the rule that interprets it.</summary>
internal sealed record IrLineSpacing(int ValueTwips, IrLineSpacingRule Rule);

/// <summary>Break kind (`w:br/@w:type`): line, page, or column.</summary>
internal enum IrBreakKind { Line, Page, Column }

/// <summary>Note kind: footnote or endnote.</summary>
internal enum IrNoteKind { Footnote, Endnote }

/// <summary>Cell vertical-merge state (`w:vMerge`): none, restart, or continue.</summary>
internal enum IrVMerge { None, Restart, Continue }

/// <summary>
/// Direct section properties (`w:sectPr`): page size, margins, orientation, and type.
/// Any unmodeled `w:sectPr` child is folded into <see cref="UnmodeledDigest"/>.
/// </summary>
internal sealed record IrSectionFormat
{
    public int? PageWidthTwips { get; init; }
    public int? PageHeightTwips { get; init; }
    public bool? Landscape { get; init; }
    public int? MarginTopTwips { get; init; }
    public int? MarginBottomTwips { get; init; }
    public int? MarginLeftTwips { get; init; }
    public int? MarginRightTwips { get; init; }

    /// <summary>`w:type/@w:val` as written (e.g. "nextPage", "continuous").</summary>
    public string? SectionType { get; init; }

    /// <summary>
    /// Digest of unmodeled `w:sectPr` children (§6.4) so a change in an unmodeled section
    /// property still flips the format fingerprint instead of being silently treated as equal.
    /// </summary>
    public required IrHash UnmodeledDigest { get; init; }
}

/// <summary>
/// Direct run properties (`w:rPr`), the v1 modeled subset. Values are stored exactly as
/// written; theme/style indirection is resolved only in effective formats (M1.3).
/// </summary>
internal sealed record IrRunFormat
{
    public string? StyleId { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public IrUnderline? Underline { get; init; }
    public bool? Strike { get; init; }
    public bool? DoubleStrike { get; init; }
    public IrVertAlign? VertAlign { get; init; }

    /// <summary>`w:rFonts/@w:ascii` as written; theme-resolved only in effective formats (M1.3).</summary>
    public string? FontAscii { get; init; }
    public int? SizeHalfPoints { get; init; }

    /// <summary>`w:color/@w:val` as written, including the literal string "auto".</summary>
    public string? ColorHex { get; init; }
    public string? Highlight { get; init; }
    public bool? Caps { get; init; }
    public bool? SmallCaps { get; init; }
    public bool? Vanish { get; init; }

    /// <summary>
    /// Digest of unmodeled `w:rPr` children (§6.4) so a change in an unmodeled run property
    /// still flips the format fingerprint instead of being silently treated as equal.
    /// </summary>
    public required IrHash UnmodeledDigest { get; init; }
}

/// <summary>
/// Direct paragraph properties (`w:pPr`), the v1 modeled subset. Values are stored exactly as
/// written; theme/style indirection is resolved only in effective formats (M1.3).
/// </summary>
internal sealed record IrParaFormat
{
    public string? StyleId { get; init; }
    public IrJustification? Justification { get; init; }
    public int? IndentLeftTwips { get; init; }
    public int? IndentRightTwips { get; init; }

    /// <summary>First-line indent in twips; a negative value means a hanging indent.</summary>
    public int? IndentFirstLineTwips { get; init; }
    public int? SpacingBeforeTwips { get; init; }
    public int? SpacingAfterTwips { get; init; }
    public IrLineSpacing? LineSpacing { get; init; }
    public int? OutlineLevel { get; init; }
    public bool? KeepNext { get; init; }
    public bool? KeepLines { get; init; }
    public bool? PageBreakBefore { get; init; }

    /// <summary>Direct list membership (`w:numPr/w:numId/@w:val`) as written; null when absent.</summary>
    public int? NumId { get; init; }

    /// <summary>Direct list level (`w:numPr/w:ilvl/@w:val`) as written; null when absent.</summary>
    public int? Ilvl { get; init; }

    /// <summary>
    /// Digest of unmodeled `w:pPr` children (§6.4) so a change in an unmodeled paragraph
    /// property still flips the format fingerprint instead of being silently treated as equal.
    /// </summary>
    public required IrHash UnmodeledDigest { get; init; }
}

/// <summary>
/// List membership facts for a paragraph, mirroring what `GetBlockMetadata`/`GetListMembership`
/// report. <see cref="AbstractNumId"/>, <see cref="NumberFormat"/>, <see cref="StartOverride"/>,
/// and <see cref="FromStyle"/> are populated in M1.3; the M1.1 reader passes the placeholders
/// null / "" / null / false respectively. A null <see cref="AbstractNumId"/> means "not yet
/// resolved" (numbering resolution lands in M1.3); never a -1 sentinel.
/// </summary>
internal sealed record IrListInfo(int NumId, int? AbstractNumId, int Ilvl,
                                  string NumberFormat, int? StartOverride, bool FromStyle);
