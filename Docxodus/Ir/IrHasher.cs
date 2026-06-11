#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// Canonicalization and hashing for the Document IR (spec §6). Produces the deterministic
/// content hashes, format fingerprints, and opaque canonical hashes the diff engine relies
/// on. Everything here is pure and deterministic: invariant-culture numeric formatting, no
/// DateTime, no randomness, big-endian digest ordering matching <see cref="IrHash.ToHex"/>.
/// </summary>
internal static class IrHasher
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // PowerTools bookkeeping namespaces (Unid/StyleName/etc.) — declared via the "pt14"
    // prefix in the codebase. Stripped from canonical form so it survives Unid churn.
    private static readonly string PtNamespace = "http://powertools.codeplex.com/2011";
    private static readonly string PtInsertNamespace = "http://powertools.codeplex.com/documentbuilder/2011/insert";

    private static readonly XName ProofErr = W + "proofErr";
    private static readonly XName NoProof = W + "noProof";

    // --- 6.3 XML canonicalization -----------------------------------------

    /// <summary>
    /// Produce canonical UTF-8 bytes for <paramref name="element"/> (spec §6.3). Deep-clones
    /// the element, then for every element in the clone: removes rsid* attributes (any
    /// namespace), PowerTools (pt14) attributes, and xmlns declarations; drops descendant
    /// <c>w:proofErr</c>/<c>w:noProof</c> elements; and sorts remaining attributes by
    /// (namespace, local name) ordinal. Serialized with <see cref="SaveOptions.DisableFormatting"/>.
    /// Result is stable across attribute reordering and rsid/Unid churn.
    /// </summary>
    /// <remarks>
    /// Deferred §6.3 clauses (intentional for M1.1):
    /// <list type="bullet">
    /// <item>"Normalize inter-element whitespace" is handled solely by
    /// <see cref="SaveOptions.DisableFormatting"/> — we do not emit indentation/line breaks
    /// between elements. Pre-existing <em>significant</em> whitespace text nodes (e.g.
    /// <c>xml:space="preserve"</c> content) are deliberately preserved as content, not
    /// collapsed.</item>
    /// <item>The broader N1/N2-rule attribute stripping (beyond the rsid*/pt14/proofErr/noProof
    /// noise removed here) lands with the M1.2 normalization work; this method intentionally
    /// strips only that minimal noise set for now.</item>
    /// </list>
    /// </remarks>
    public static byte[] Canonicalize(XElement element)
    {
        var clone = new XElement(element);
        Clean(clone);
        var xml = clone.ToString(SaveOptions.DisableFormatting);
        return Encoding.UTF8.GetBytes(xml);
    }

    /// <summary>SHA-256 of <see cref="Canonicalize(XElement)"/>.</summary>
    public static IrHash CanonicalHash(XElement element) => IrHash.Compute(Canonicalize(element));

    private static void Clean(XElement element)
    {
        // Remove noise child elements first (proofErr/noProof, anywhere in the subtree).
        var toRemove = element
            .Descendants()
            .Where(d => d.Name == ProofErr || d.Name == NoProof)
            .ToList();
        foreach (var d in toRemove)
            d.Remove();

        foreach (var el in element.DescendantsAndSelf())
            CleanAttributes(el);
    }

    private static void CleanAttributes(XElement element)
    {
        var kept = element.Attributes()
            .Where(a => !ShouldStripAttribute(a))
            .OrderBy(a => a.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(a => a.Name.LocalName, StringComparer.Ordinal)
            .ToList();

        element.RemoveAttributes();
        foreach (var a in kept)
            element.Add(a);
    }

    private static bool ShouldStripAttribute(XAttribute attribute)
    {
        // xmlns namespace-declaration attributes.
        if (attribute.IsNamespaceDeclaration)
            return true;

        var ns = attribute.Name.NamespaceName;

        // PowerTools (pt14) bookkeeping attributes — any in those namespaces.
        if (ns == PtNamespace || ns == PtInsertNamespace)
            return true;

        // rsid* attributes in any namespace (local name starts with "rsid").
        if (attribute.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal))
            return true;

        return false;
    }

    // --- 6.2 Format fingerprints ------------------------------------------

    /// <summary>
    /// Fingerprint of a direct run format (spec §6.2). Serializes every NON-NULL property as
    /// length-prefixed <c>name=&lt;charCount&gt;:&lt;value&gt;;</c> pairs in declaration order
    /// (the char count is the value's <see cref="string.Length"/> in invariant culture),
    /// appends the 32 raw <see cref="IrRunFormat.UnmodeledDigest"/> bytes, then SHA-256s the
    /// whole. The length prefix makes the framing unambiguous even when a value (e.g. a style
    /// id or font name) contains <c>=</c> or <c>;</c>. Null properties are omitted entirely,
    /// so "Bold=true, Italic=null" hashes equal to "Bold=true" with Italic absent.
    /// </summary>
    public static IrHash FingerprintRunFormat(IrRunFormat f)
    {
        var sb = new StringBuilder();
        AppendField(sb, "StyleId", f.StyleId);
        AppendField(sb, "Bold", f.Bold);
        AppendField(sb, "Italic", f.Italic);
        AppendField(sb, "Underline", RenderUnderline(f.Underline));
        AppendField(sb, "Strike", f.Strike);
        AppendField(sb, "DoubleStrike", f.DoubleStrike);
        AppendField(sb, "VertAlign", f.VertAlign?.ToString());
        AppendField(sb, "FontAscii", f.FontAscii);
        AppendField(sb, "SizeHalfPoints", f.SizeHalfPoints);
        AppendField(sb, "ColorHex", f.ColorHex);
        AppendField(sb, "Highlight", f.Highlight);
        AppendField(sb, "Caps", f.Caps);
        AppendField(sb, "SmallCaps", f.SmallCaps);
        AppendField(sb, "Vanish", f.Vanish);
        return HashFields(sb, f.UnmodeledDigest);
    }

    /// <summary>Fingerprint of a direct paragraph format (spec §6.2). See <see cref="FingerprintRunFormat"/>.</summary>
    public static IrHash FingerprintParaFormat(IrParaFormat f)
    {
        var sb = new StringBuilder();
        AppendField(sb, "StyleId", f.StyleId);
        AppendField(sb, "Justification", f.Justification?.ToString());
        AppendField(sb, "IndentLeftTwips", f.IndentLeftTwips);
        AppendField(sb, "IndentRightTwips", f.IndentRightTwips);
        AppendField(sb, "IndentFirstLineTwips", f.IndentFirstLineTwips);
        AppendField(sb, "SpacingBeforeTwips", f.SpacingBeforeTwips);
        AppendField(sb, "SpacingAfterTwips", f.SpacingAfterTwips);
        AppendField(sb, "LineSpacing", RenderLineSpacing(f.LineSpacing));
        AppendField(sb, "OutlineLevel", f.OutlineLevel);
        AppendField(sb, "KeepNext", f.KeepNext);
        AppendField(sb, "KeepLines", f.KeepLines);
        AppendField(sb, "PageBreakBefore", f.PageBreakBefore);
        return HashFields(sb, f.UnmodeledDigest);
    }

    /// <summary>Fingerprint of a direct section format (spec §6.2). See <see cref="FingerprintRunFormat"/>.</summary>
    public static IrHash FingerprintSectionFormat(IrSectionFormat f)
    {
        var sb = new StringBuilder();
        AppendField(sb, "PageWidthTwips", f.PageWidthTwips);
        AppendField(sb, "PageHeightTwips", f.PageHeightTwips);
        AppendField(sb, "Landscape", f.Landscape);
        AppendField(sb, "MarginTopTwips", f.MarginTopTwips);
        AppendField(sb, "MarginBottomTwips", f.MarginBottomTwips);
        AppendField(sb, "MarginLeftTwips", f.MarginLeftTwips);
        AppendField(sb, "MarginRightTwips", f.MarginRightTwips);
        AppendField(sb, "SectionType", f.SectionType);
        return HashFields(sb, f.UnmodeledDigest);
    }

    /// <summary>
    /// Block-level format fingerprint (spec §6.2): the paragraph fingerprint's 32 bytes
    /// followed by each run fingerprint's 32 bytes in order, SHA-256'd. Run order is
    /// significant, so a single bolded word flips the block fingerprint.
    /// </summary>
    public static IrHash FingerprintBlock(IrParaFormat paraFormat, IEnumerable<IrRunFormat> runFormats)
    {
        using var ms = new MemoryStream();
        WriteHash(ms, FingerprintParaFormat(paraFormat));
        foreach (var rf in runFormats)
            WriteHash(ms, FingerprintRunFormat(rf));
        return IrHash.Compute(ms.GetBuffer().AsSpan(0, (int)ms.Length));
    }

    /// <summary>
    /// Stable rendering of a nested <see cref="IrUnderline"/>: <c>Kind|ColorHex</c>, with the
    /// color omitted (no trailing pipe) when null — e.g. <c>Single|FF0000</c> or <c>Single</c>.
    /// </summary>
    private static string? RenderUnderline(IrUnderline? u)
    {
        if (u is null)
            return null;
        return u.ColorHex is null ? u.Kind.ToString() : $"{u.Kind}|{u.ColorHex}";
    }

    /// <summary>
    /// Stable rendering of a nested <see cref="IrLineSpacing"/>: <c>ValueTwips|Rule</c>
    /// (invariant culture), e.g. <c>240|Auto</c>.
    /// </summary>
    private static string? RenderLineSpacing(IrLineSpacing? ls)
    {
        if (ls is null)
            return null;
        return $"{ls.ValueTwips.ToString(CultureInfo.InvariantCulture)}|{ls.Rule}";
    }

    private static void AppendField(StringBuilder sb, string name, string? value)
    {
        if (value is null)
            return;
        AppendFramed(sb, name, value);
    }

    private static void AppendField(StringBuilder sb, string name, bool? value)
    {
        if (value is null)
            return;
        AppendFramed(sb, name, value.Value ? "true" : "false");
    }

    private static void AppendField(StringBuilder sb, string name, int? value)
    {
        if (value is null)
            return;
        AppendFramed(sb, name, value.Value.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Emit a length-prefixed <c>name=&lt;charCount&gt;:&lt;value&gt;;</c> field.
    /// <paramref name="value"/> may originate from OOXML strings (style ids, font names,
    /// color hexes, highlight/section-type tokens) that can themselves contain <c>=</c> or
    /// <c>;</c>; the leading character count (value length, invariant culture) makes the
    /// framing unambiguous so e.g. <c>StyleId="A;Bold=true"</c> can never collide with
    /// <c>StyleId="A"</c> + <c>Bold=true</c>.
    /// </summary>
    private static void AppendFramed(StringBuilder sb, string name, string value)
    {
        sb.Append(name)
          .Append('=')
          .Append(value.Length.ToString(CultureInfo.InvariantCulture))
          .Append(':')
          .Append(value)
          .Append(';');
    }

    private static IrHash HashFields(StringBuilder sb, IrHash unmodeledDigest)
    {
        var fieldBytes = Encoding.UTF8.GetBytes(sb.ToString());
        var buffer = new byte[fieldBytes.Length + 32];
        Array.Copy(fieldBytes, buffer, fieldBytes.Length);
        unmodeledDigest.CopyTo(buffer.AsSpan(fieldBytes.Length, 32));
        return IrHash.Compute(buffer);
    }

    private static void WriteHash(Stream stream, IrHash hash)
    {
        Span<byte> bytes = stackalloc byte[32];
        hash.CopyTo(bytes);
        stream.Write(bytes);
    }
}

/// <summary>
/// Accumulates the canonical byte stream for a block's <c>ContentHash</c> (spec §6.1) and
/// SHA-256s it on <see cref="Build"/>. Text is appended as UTF-8; non-text structure is
/// appended as sentinel/marker byte sequences that cannot collide with text.
/// </summary>
/// <remarks>
/// SAFETY: sentinel and structure lead bytes are <c>0x01</c> and <c>0x02</c>, both of which
/// XML 1.0 forbids in element content (U+0001 / U+0002 are not legal XML characters). No
/// XML-sourced text can therefore contain them, so a text run can never be mistaken for a
/// sentinel/marker or vice versa.
/// </remarks>
internal sealed class IrContentHashBuilder
{
    // Sentinel kinds (written after a 0x01 lead byte) — non-text inlines.
    public const byte SentinelTab = 0x01;
    public const byte SentinelLineBreak = 0x02;
    public const byte SentinelPageBreak = 0x03;
    public const byte SentinelColumnBreak = 0x04;
    public const byte SentinelFootnoteRef = 0x05;
    public const byte SentinelEndnoteRef = 0x06;
    public const byte SentinelImage = 0x07;

    // Hyperlink framing (spec §6.1, N14): the target string is bracketed between these two
    // sentinels so linked text never collides with identical plain text and a target change is
    // a content change. The child inlines' bytes follow after the closing sentinel.
    public const byte SentinelHyperlink = 0x08;
    public const byte SentinelHyperlinkTargetEnd = 0x09;

    // Textbox framing (spec §6.1 M1.5 addendum): a w:txbxContent body. The sentinel is followed by
    // each inner block's 32-byte ContentHash in order (AppendHash), so a textbox's text participates
    // in the containing paragraph's ContentHash — closing the blind spot where textbox text was
    // invisible to the diff engine — yet stays distinguishable from identical inline (non-textbox)
    // text because the sentinel and per-block hash framing can never be reproduced by literal runs.
    public const byte SentinelTextbox = 0x0B;

    public const byte SentinelOpaque = 0x0F;

    // Structure markers (written after a 0x02 lead byte) — table structure.
    public const byte StructureRow = 0x10;
    public const byte StructureCell = 0x11;

    private const byte SentinelLead = 0x01;
    private const byte StructureLead = 0x02;

    private readonly MemoryStream _buffer = new();

    /// <summary>
    /// Append the UTF-8 bytes of <paramref name="text"/>. Safe against sentinel collision
    /// because XML 1.0 forbids the sentinel/structure lead code points (U+0001/U+0002) in
    /// content, so no XML-sourced string can contain them (see type remarks).
    /// </summary>
    public void AppendText(string text)
    {
        var bytes = Encoding.UTF8.GetBytes(text);
        _buffer.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Append a non-text inline sentinel: <c>0x01</c> then <paramref name="kind"/>.</summary>
    public void AppendSentinel(byte kind)
    {
        _buffer.WriteByte(SentinelLead);
        _buffer.WriteByte(kind);
    }

    /// <summary>Append a table-structure marker: <c>0x02</c> then <paramref name="marker"/>.</summary>
    public void AppendStructure(byte marker)
    {
        _buffer.WriteByte(StructureLead);
        _buffer.WriteByte(marker);
    }

    /// <summary>Append the 32 raw bytes of <paramref name="hash"/> (e.g. an image or opaque hash).</summary>
    public void AppendHash(IrHash hash)
    {
        Span<byte> bytes = stackalloc byte[32];
        hash.CopyTo(bytes);
        _buffer.Write(bytes);
    }

    /// <summary>SHA-256 of the accumulated byte stream.</summary>
    public IrHash Build() => IrHash.Compute(_buffer.GetBuffer().AsSpan(0, (int)_buffer.Length));
}
