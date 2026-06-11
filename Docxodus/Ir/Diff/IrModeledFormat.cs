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
}
