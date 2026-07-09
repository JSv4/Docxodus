#nullable enable

using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Diff-time MODELED-ONLY format projection (M2.2 Task 4). Produces a string key for an
/// <see cref="IrRunFormat"/> that includes every modeled field but EXCLUDES
/// <see cref="IrRunFormat.UnmodeledDigest"/>, and a boundary-normalized modeled-only block signature
/// for a paragraph. Both are purely diff-time — the IR's stored hashes are untouched.
/// </summary>
/// <remarks>
/// <para><b>Why a string key and not the record.</b> Record equality on <see cref="IrRunFormat"/>
/// folds in <see cref="IrRunFormat.UnmodeledDigest"/>, which is exactly the noise channel
/// (<c>w:lang</c>/<c>w:bCs</c>/<c>w:iCs</c>/…) the WC-BodyBookmarks diagnosis flagged. The key below
/// enumerates the modeled fields ONLY, so two runs that differ solely in unmodeled rPr children produce
/// the same key. Field framing mirrors <see cref="IrHasher.FingerprintRunFormat"/> (name + value, null
/// fields omitted) minus the trailing digest.</para>
/// <para><b>Boundary normalization (block level).</b> <see cref="BlockSignature"/> walks the paragraph's
/// DIFF TOKENS (not its raw runs) and emits one <c>(MatchKey, modeled-format key)</c> pair per token.
/// Because the token stream is run-boundary-independent — a word split across two runs on one side and
/// one run on the other tokenizes to the SAME token sequence — the signature is invariant to the
/// run-resegmentation churn that flips the reader's stored block FormatFingerprint (the M2.1 finding).
/// Two ContentHash-equal paragraphs therefore compare format-equal iff their per-token MODELED formats
/// agree, regardless of how editing churned the run boundaries.</para>
/// </remarks>
internal static class IrModeledFormat
{
    /// <summary>
    /// Modeled-only equality key for a run format: every modeled field, framed; the unmodeled digest is
    /// deliberately omitted. A null format maps to the empty key (consistent with a run carrying no rPr).
    /// </summary>
    public static string RunKey(IrRunFormat? f)
    {
        if (f is null)
            return string.Empty;

        var sb = new StringBuilder();
        Append(sb, "StyleId", f.StyleId);
        Append(sb, "Bold", f.Bold);
        Append(sb, "Italic", f.Italic);
        Append(sb, "Underline", RenderUnderline(f.Underline));
        Append(sb, "Strike", f.Strike);
        Append(sb, "DoubleStrike", f.DoubleStrike);
        Append(sb, "VertAlign", f.VertAlign?.ToString());
        Append(sb, "FontAscii", f.FontAscii);
        Append(sb, "SizeHalfPoints", f.SizeHalfPoints);
        Append(sb, "ColorHex", f.ColorHex);
        Append(sb, "Highlight", f.Highlight);
        Append(sb, "Caps", f.Caps);
        Append(sb, "SmallCaps", f.SmallCaps);
        Append(sb, "Vanish", f.Vanish);
        return sb.ToString();
    }

    /// <summary>
    /// Modeled-only equality key for a PARAGRAPH format (block-format-change family, 2026-07-03):
    /// every modeled <see cref="IrParaFormat"/> field, framed like <see cref="RunKey"/>; the unmodeled
    /// digest is deliberately omitted. A null format maps to the empty key (a paragraph with no pPr).
    /// </summary>
    public static string ParaKey(IrParaFormat? f)
    {
        if (f is null)
            return string.Empty;

        var sb = new StringBuilder();
        Append(sb, "PStyleId", f.StyleId);
        Append(sb, "Jc", f.Justification?.ToString());
        Append(sb, "IndL", f.IndentLeftTwips);
        Append(sb, "IndR", f.IndentRightTwips);
        Append(sb, "IndFL", f.IndentFirstLineTwips);
        Append(sb, "SpB", f.SpacingBeforeTwips);
        Append(sb, "SpA", f.SpacingAfterTwips);
        Append(sb, "Line", RenderLineSpacing(f.LineSpacing));
        Append(sb, "Outline", f.OutlineLevel);
        Append(sb, "KeepNext", f.KeepNext);
        Append(sb, "KeepLines", f.KeepLines);
        Append(sb, "PBB", f.PageBreakBefore);
        Append(sb, "NumId", f.NumId);
        Append(sb, "Ilvl", f.Ilvl);
        return sb.ToString();
    }

    /// <summary>
    /// True iff two run formats are equal for diff purposes under <paramref name="comparison"/>:
    /// modeled-field equality (ignoring the unmodeled digest) for
    /// <see cref="IrFormatComparison.ModeledOnly"/>, full record equality for
    /// <see cref="IrFormatComparison.Full"/>.
    /// </summary>
    public static bool RunFormatEqual(IrRunFormat? a, IrRunFormat? b, IrFormatComparison comparison)
    {
        if (comparison == IrFormatComparison.Full)
            return EqualityComparer<IrRunFormat?>.Default.Equals(a, b);
        return RunKey(a) == RunKey(b);
    }

    /// <summary>
    /// Boundary-normalized modeled-only block signature of a paragraph: the concatenation of one
    /// <c>«MatchKey␟modeled-format-key␞»</c> record per diff token. Two paragraphs with equal signatures
    /// have the same text AND the same per-token modeled formatting, independent of run boundaries.
    /// </summary>
    public static string BlockSignature(IrParagraph paragraph, IrDiffSettings settings)
    {
        var tokens = IrDiffTokenizer.Tokenize(paragraph, settings);
        var sb = new StringBuilder();
        foreach (var t in tokens)
        {
            sb.Append(t.MatchKey);
            sb.Append('␟'); // unit separator glyph (not an XML-legal content char source)
            sb.Append(RunKey(t.Format));
            sb.Append('␞'); // record separator glyph
        }

        // Block-format-change family (2026-07-03): the paragraph's own modeled format participates in
        // the signature, so a pPr-only change (jc/indent/spacing/style/numbering) classifies FormatOnly
        // instead of Unchanged. Gated on the PARAGRAPH slice so the composite can turn pPr merge ON (B1)
        // while keeping section OFF (B2) — in two-way the two flags are equal, so behavior is identical.
        if (settings.TrackParagraphFormatChanges)
        {
            sb.Append('¶');
            sb.Append(ParaKey(paragraph.Format));
        }
        // A3: an inline (in-pPr) sectPr's modeled page setup participates too, so a mid-document
        // sectPr-only change classifies FormatOnly instead of Unchanged under ModeledOnly. Gated on the
        // SECTION slice so the composite can turn inline-section merge ON (B2) independently of paragraph/table.
        if (settings.TrackSectionFormatChanges)
        {
            sb.Append('§');
            sb.Append(SectionKey(paragraph.InlineSectionFormat));
        }

        return sb.ToString();
    }

    /// <summary>Modeled-only equality key for a SECTION format (block-format follow-up A3): the modeled
    /// <see cref="IrSectionFormat"/> fields, framed like <see cref="ParaKey"/>; null maps to the empty key.</summary>
    public static string SectionKey(IrSectionFormat? f)
    {
        if (f is null)
            return string.Empty;
        var sb = new StringBuilder();
        Append(sb, "PgW", f.PageWidthTwips);
        Append(sb, "PgH", f.PageHeightTwips);
        Append(sb, "Land", f.Landscape);
        Append(sb, "MT", f.MarginTopTwips);
        Append(sb, "MB", f.MarginBottomTwips);
        Append(sb, "ML", f.MarginLeftTwips);
        Append(sb, "MR", f.MarginRightTwips);
        Append(sb, "SType", f.SectionType);
        return sb.ToString();
    }

    /// <summary>
    /// Project a run format's MODELED fields to a WmlComparer-friendly property dictionary (M2.3 Task 1):
    /// name → display value, omitting any field that is null on this run (mirroring WmlComparer's
    /// "absent rPr child ⇒ absent key" convention). Property names match
    /// <c>WmlComparer.GetFriendlyPropertyName</c> (bold/italic/underline/strikethrough/…) so the produced
    /// <see cref="IrFormatChangeDetails"/> is adapter-comparable to <c>WmlComparer.FormatChangeDetails</c>.
    /// The unmodeled digest is never projected (it is undescribable as an rPrChange).
    /// </summary>
    public static IReadOnlyDictionary<string, string> ModeledProperties(IrRunFormat? f)
    {
        var dict = new Dictionary<string, string>();
        if (f is null)
            return dict;

        AddProp(dict, "style", f.StyleId);
        AddBool(dict, "bold", f.Bold);
        AddBool(dict, "italic", f.Italic);
        AddProp(dict, "underline", RenderUnderline(f.Underline));
        AddBool(dict, "strikethrough", f.Strike);
        AddBool(dict, "doubleStrikethrough", f.DoubleStrike);
        AddProp(dict, "verticalAlign", f.VertAlign?.ToString());
        AddProp(dict, "font", f.FontAscii);
        AddInt(dict, "fontSize", f.SizeHalfPoints);
        AddProp(dict, "color", f.ColorHex);
        AddProp(dict, "highlight", f.Highlight);
        AddBool(dict, "allCaps", f.Caps);
        AddBool(dict, "smallCaps", f.SmallCaps);
        AddBool(dict, "hidden", f.Vanish);
        return dict;
    }

    /// <summary>
    /// Build the <see cref="IrFormatChangeDetails"/> for a (left, right) run-format pair: the modeled
    /// property dictionaries plus the changed-property names (a field present-on-one-side-only OR
    /// present-on-both-with-differing-value), computed by the SAME rule as
    /// <c>WmlComparer.ExtractFormatChangeDetails</c>. Changed names are emitted in a STABLE order (the
    /// fixed modeled-field order of <see cref="ModeledProperties"/>) for deterministic output.
    /// </summary>
    public static IrFormatChangeDetails FormatChangeDetails(IrRunFormat? left, IrRunFormat? right)
    {
        var oldProps = ModeledProperties(left);
        var newProps = ModeledProperties(right);

        var changed = new List<string>();
        foreach (var name in ModeledFieldOrder)
        {
            bool hasOld = oldProps.TryGetValue(name, out var oldVal);
            bool hasNew = newProps.TryGetValue(name, out var newVal);
            if (hasOld != hasNew || (hasOld && hasNew && oldVal != newVal))
                changed.Add(name);
        }

        return new IrFormatChangeDetails(oldProps, newProps, changed);
    }

    /// <summary>The fixed modeled-field property-name order (matches <see cref="ModeledProperties"/>).</summary>
    private static readonly string[] ModeledFieldOrder =
    {
        "style", "bold", "italic", "underline", "strikethrough", "doubleStrikethrough",
        "verticalAlign", "font", "fontSize", "color", "highlight", "allCaps", "smallCaps", "hidden",
    };

    /// <summary>
    /// Project a PARAGRAPH format's modeled fields to a property dictionary (block-format-change family):
    /// name → display value, omitting null fields — the paragraph analogue of
    /// <see cref="ModeledProperties"/>. The unmodeled digest is never projected (undescribable).
    /// </summary>
    public static IReadOnlyDictionary<string, string> ModeledParaProperties(IrParaFormat? f)
    {
        var dict = new Dictionary<string, string>();
        if (f is null)
            return dict;

        AddProp(dict, "style", f.StyleId);
        AddProp(dict, "justification", f.Justification?.ToString());
        AddInt(dict, "indentLeft", f.IndentLeftTwips);
        AddInt(dict, "indentRight", f.IndentRightTwips);
        AddInt(dict, "indentFirstLine", f.IndentFirstLineTwips);
        AddInt(dict, "spacingBefore", f.SpacingBeforeTwips);
        AddInt(dict, "spacingAfter", f.SpacingAfterTwips);
        AddProp(dict, "lineSpacing", RenderLineSpacing(f.LineSpacing));
        AddInt(dict, "outlineLevel", f.OutlineLevel);
        AddBool(dict, "keepNext", f.KeepNext);
        AddBool(dict, "keepLines", f.KeepLines);
        AddBool(dict, "pageBreakBefore", f.PageBreakBefore);
        AddInt(dict, "numId", f.NumId);
        AddInt(dict, "numLevel", f.Ilvl);
        return dict;
    }

    /// <summary>
    /// Build the Paragraph-scope <see cref="IrFormatChangeDetails"/> for a (left, right) paragraph-format
    /// pair: modeled property dictionaries + changed names in the fixed field order — the same
    /// present-on-one-side-or-differing rule as <see cref="FormatChangeDetails"/>. An unmodeled-only delta
    /// (detected under Full) yields empty changed names, mirroring the run-level empty-details convention.
    /// </summary>
    public static IrFormatChangeDetails ParaFormatChangeDetails(IrParaFormat? left, IrParaFormat? right)
    {
        var oldProps = ModeledParaProperties(left);
        var newProps = ModeledParaProperties(right);

        var changed = new List<string>();
        foreach (var name in ModeledParaFieldOrder)
        {
            bool hasOld = oldProps.TryGetValue(name, out var oldVal);
            bool hasNew = newProps.TryGetValue(name, out var newVal);
            if (hasOld != hasNew || (hasOld && hasNew && oldVal != newVal))
                changed.Add(name);
        }

        return new IrFormatChangeDetails(oldProps, newProps, changed, IrFormatChangeScope.Paragraph);
    }

    /// <summary>The fixed modeled paragraph-field property-name order (matches <see cref="ModeledParaProperties"/>).</summary>
    private static readonly string[] ModeledParaFieldOrder =
    {
        "style", "justification", "indentLeft", "indentRight", "indentFirstLine",
        "spacingBefore", "spacingAfter", "lineSpacing", "outlineLevel",
        "keepNext", "keepLines", "pageBreakBefore", "numId", "numLevel",
    };

    /// <summary>Project a SECTION format's modeled fields to a property dictionary (block-format-change family,
    /// Phase 3): the section analogue of <see cref="ModeledParaProperties"/>; null fields omitted.</summary>
    public static IReadOnlyDictionary<string, string> ModeledSectionProperties(IrSectionFormat? f)
    {
        var dict = new Dictionary<string, string>();
        if (f is null)
            return dict;

        AddInt(dict, "pageWidth", f.PageWidthTwips);
        AddInt(dict, "pageHeight", f.PageHeightTwips);
        AddBool(dict, "landscape", f.Landscape);
        AddInt(dict, "marginTop", f.MarginTopTwips);
        AddInt(dict, "marginBottom", f.MarginBottomTwips);
        AddInt(dict, "marginLeft", f.MarginLeftTwips);
        AddInt(dict, "marginRight", f.MarginRightTwips);
        AddProp(dict, "sectionType", f.SectionType);
        return dict;
    }

    /// <summary>Build the Section-scope <see cref="IrFormatChangeDetails"/> for a (left, right) section-format
    /// pair — modeled fields only (page setup/margins/orientation/type). An unmodeled-only section change
    /// (e.g. <c>w:cols</c>) yields empty changed names: it is tracked by the markup (canonical) but not
    /// described here — the modeled-only-revision limitation, consistent with the run/paragraph scopes.</summary>
    public static IrFormatChangeDetails SectionFormatChangeDetails(IrSectionFormat? left, IrSectionFormat? right)
    {
        var oldProps = ModeledSectionProperties(left);
        var newProps = ModeledSectionProperties(right);

        var changed = new List<string>();
        foreach (var name in ModeledSectionFieldOrder)
        {
            bool hasOld = oldProps.TryGetValue(name, out var oldVal);
            bool hasNew = newProps.TryGetValue(name, out var newVal);
            if (hasOld != hasNew || (hasOld && hasNew && oldVal != newVal))
                changed.Add(name);
        }

        return new IrFormatChangeDetails(oldProps, newProps, changed, IrFormatChangeScope.Section);
    }

    private static readonly string[] ModeledSectionFieldOrder =
    {
        "pageWidth", "pageHeight", "landscape", "marginTop", "marginBottom",
        "marginLeft", "marginRight", "sectionType",
    };

    private static void AddProp(Dictionary<string, string> dict, string name, string? value)
    {
        if (value is not null)
            dict[name] = value;
    }

    private static void AddBool(Dictionary<string, string> dict, string name, bool? value)
    {
        if (value is not null)
            dict[name] = value.Value ? "true" : "false";
    }

    private static void AddInt(Dictionary<string, string> dict, string name, int? value)
    {
        if (value is not null)
            dict[name] = value.Value.ToString(CultureInfo.InvariantCulture);
    }

    // ------------------------------------------------------------------ framing

    private static void Append(StringBuilder sb, string name, string? value)
    {
        if (value is null)
            return;
        sb.Append(name).Append('=')
          .Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':')
          .Append(value).Append(';');
    }

    private static void Append(StringBuilder sb, string name, bool? value)
    {
        if (value is not null)
            Append(sb, name, value.Value ? "true" : "false");
    }

    private static void Append(StringBuilder sb, string name, int? value)
    {
        if (value is not null)
            Append(sb, name, value.Value.ToString(CultureInfo.InvariantCulture));
    }

    private static string? RenderUnderline(IrUnderline? u)
    {
        if (u is null)
            return null;
        return u.ColorHex is null ? u.Kind.ToString() : $"{u.Kind}|{u.ColorHex}";
    }

    private static string? RenderLineSpacing(IrLineSpacing? ls)
    {
        if (ls is null)
            return null;
        return $"{ls.ValueTwips.ToString(CultureInfo.InvariantCulture)}|{ls.Rule}";
    }
}
