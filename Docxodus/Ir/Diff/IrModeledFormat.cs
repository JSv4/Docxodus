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
