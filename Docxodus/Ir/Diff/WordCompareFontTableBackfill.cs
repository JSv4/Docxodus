#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Backfills <c>word/fontTable.xml</c> and <c>word/webSettings.xml</c> into a DocxDiff output, matching
/// Word's compare behavior — Word SYNTHESIZES both parts for every output document even when the inputs
/// carry neither (the source documents rarely carry them).
/// <para><b>Why it matters (a real fidelity fix, not a nicety).</b> A fontTable declares each font's
/// <c>panose1/charset/family/pitch</c> metrics, which is what LibreOffice consults to pick a substitute
/// for a font it does not have installed (Aptos, Calibri, or a raw CSS stack). When Word's redline carries a
/// fontTable and our output does not, the two documents substitute the SAME absent font DIFFERENTLY — so
/// even a byte-identical body renders to different glyphs/metrics, and a rendered redline diverges from
/// Word's compare output. Supplying a matching fontTable makes both sides substitute identically. The
/// theme/docDefaults backfills are the analogous single-owner passes; this completes the set. webSettings
/// is emitted empty-bodied (as Word's is for these documents) purely so the package parity holds.</para>
/// </summary>
internal static class WordCompareFontTableBackfill
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Fonts Word's fontTable always carries (its stock theme + default faces), in Word's order.</summary>
    private static readonly string[] StockFonts = { "Times New Roman", "Aptos", "Calibri", "Aptos Display" };

    /// <summary>Exact <c>&lt;w:font&gt;</c> inner XML for the fonts seen in Word's compare output (panose/charset/
    /// family/pitch/sig verbatim from Word). Fonts not listed get a generic swiss/variable descriptor
    /// (a sans-serif substitute hint) — enough for LibreOffice's substitution to match Word's.</summary>
    private static readonly Dictionary<string, string> KnownFontInner = new()
    {
        ["Times New Roman"] = "<w:panose1 w:val=\"02020603050405020304\"/><w:charset w:val=\"00\"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E0002EFF\" w:usb1=\"C000785B\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/>",
        ["Calibri"] = "<w:panose1 w:val=\"020F0502020204030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E0002AFF\" w:usb1=\"C000247B\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/>",
        ["Aptos"] = "<w:panose1 w:val=\"020B0004020202020204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"20000287\" w:usb1=\"00000003\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"0000019F\" w:csb1=\"00000000\"/>",
        ["Aptos Display"] = "<w:panose1 w:val=\"020B0004020202020204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"20000287\" w:usb1=\"00000003\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"0000019F\" w:csb1=\"00000000\"/>",
        ["Arial"] = "<w:panose1 w:val=\"020B0604020202020204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E0002AFF\" w:usb1=\"C0007843\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/>",
        ["Symbol"] = "<w:panose1 w:val=\"05050102010706020507\"/><w:charset w:val=\"02\"/><w:family w:val=\"decorative\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"00000000\" w:usb1=\"10000000\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"80000000\" w:csb1=\"00000000\"/>",
        ["Courier New"] = "<w:panose1 w:val=\"02070309020205020404\"/><w:charset w:val=\"00\"/><w:family w:val=\"modern\"/><w:pitch w:val=\"fixed\"/><w:sig w:usb0=\"E0002AFF\" w:usb1=\"C0007843\" w:usb2=\"00000009\" w:usb3=\"00000000\" w:csb0=\"000001FF\" w:csb1=\"00000000\"/>",
        ["Cambria"] = "<w:panose1 w:val=\"02040503050406030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"roman\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"E00002FF\" w:usb1=\"400004FF\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"0000019F\" w:csb1=\"00000000\"/>",
        ["Courier"] = "<w:panose1 w:val=\"02070309020205020404\"/><w:charset w:val=\"00\"/><w:family w:val=\"auto\"/><w:pitch w:val=\"variable\"/><w:sig w:usb0=\"00000003\" w:usb1=\"00000000\" w:usb2=\"00000000\" w:usb3=\"00000000\" w:csb0=\"00000001\" w:csb1=\"00000000\"/>",
    };

    /// <summary>Generic descriptor for an unlisted font.
    /// <para>A CSS FONT STACK (a comma-bearing value like <c>"Roboto, sans-serif"</c> that some HTML→DOCX
    /// producers write straight into <c>w:rFonts</c>, and which Word's compare keeps verbatim) is declared
    /// exactly as Word declares it: an <c>&lt;w:altName&gt;</c> giving the PRIMARY family (the first stack
    /// component, unquoted) — which is what LibreOffice actually resolves for substitution — plus Word's
    /// generic non-TrueType descriptor. Emitting a plain descriptor WITHOUT the altName makes LibreOffice
    /// substitute the raw stack string differently than Word's compare output (which carries the altName),
    /// regressing those web-font documents.</para>
    /// A plain single name that reads as serif/monospace is nudged to the matching family so a substitute
    /// is picked in kind.</summary>
    private static string GenericInner(string name)
    {
        int comma = name.IndexOf(',');
        if (comma > 0)
        {
            string primary = name.Substring(0, comma).Trim().Trim('"', '\'');
            return $"<w:altName w:val=\"{System.Security.SecurityElement.Escape(primary)}\"/>" +
                   "<w:panose1 w:val=\"020B0604020202020204\"/><w:charset w:val=\"00\"/>" +
                   "<w:family w:val=\"roman\"/><w:notTrueType/><w:pitch w:val=\"variable\"/>";
        }
        string lower = name.ToLowerInvariant();
        string family = lower.Contains("serif") && !lower.Contains("sans")
            ? "roman"
            : lower.Contains("mono") || lower.Contains("courier") || lower.Contains("consol")
                ? "modern"
                : "swiss";
        string pitch = family == "modern" ? "fixed" : "variable";
        return $"<w:charset w:val=\"00\"/><w:family w:val=\"{family}\"/><w:pitch w:val=\"{pitch}\"/>";
    }

    /// <summary>Emit fontTable.xml (if absent) and webSettings.xml (if absent) into <paramref name="main"/>,
    /// listing Word's stock fonts plus every font the document/styles actually reference. When a fontTable
    /// is CARRIED THROUGH from the source instead, normalize its known-font declarations to the same stock
    /// metadata — Word never copies an input fontTable verbatim; it regenerates every declaration from its
    /// installed-font metadata, so a degraded input declaration (family="auto", zeroed panose) must not
    /// survive into the output or LibreOffice substitutes the font differently than it does for Word's
    /// own compare output.</summary>
    public static void Backfill(MainDocumentPart main)
    {
        if (main.FontTablePart is null)
            BackfillFontTable(main);
        else
            NormalizeCarriedFontTable(main.FontTablePart);
        if (main.WebSettingsPart is null)
            BackfillWebSettings(main);
    }

    /// <summary>
    /// Rewrite the descriptor children of every KNOWN font declaration in a carried-through fontTable to
    /// Word's exact metadata (the same <see cref="KnownFontInner"/> table the from-scratch backfill uses).
    /// Unknown fonts are left byte-untouched (their declarations may be the only substitution hint a
    /// custom face has), and relationship-backed embedded-font children (<c>w:embedRegular</c> etc.) are
    /// preserved — only the metrics metadata is normalized.
    /// </summary>
    private static void NormalizeCarriedFontTable(FontTablePart part)
    {
        var xdoc = part.GetXDocument();
        var root = xdoc.Root;
        if (root is null)
            return;

        bool changed = false;
        foreach (var font in root.Elements(W + "font"))
        {
            var name = (string?)font.Attribute(W + "name");
            if (name is null || !KnownFontInner.TryGetValue(name, out var inner))
                continue;

            var embeds = font.Elements()
                .Where(e => e.Name.Namespace == W && e.Name.LocalName.StartsWith("embed", System.StringComparison.Ordinal))
                .ToList();
            var replacement = XElement.Parse(
                "<w:font xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                inner + "</w:font>");

            font.RemoveNodes();
            font.Add(replacement.Elements());
            font.Add(embeds);
            changed = true;
        }

        if (changed)
            part.PutXDocument();
    }

    private static void BackfillFontTable(MainDocumentPart main)
    {
        // Ordered, deduped font list: Word's stock faces first, then every font referenced by a
        // w:rFonts (ascii/hAnsi/cs/eastAsia) in the body or the styles part — the fonts LibreOffice
        // must be told how to substitute.
        var names = new List<string>();
        var seen = new HashSet<string>();
        void Add(string? n)
        {
            if (!string.IsNullOrEmpty(n) && seen.Add(n!))
                names.Add(n!);
        }
        foreach (var n in StockFonts)
            Add(n);
        foreach (var root in EnumerateFontSources(main))
            foreach (var rFonts in root.Descendants(W + "rFonts"))
                foreach (var attr in new[] { "ascii", "hAnsi", "cs", "eastAsia" })
                    Add((string?)rFonts.Attribute(W + attr));

        var sb = new System.Text.StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<w:fonts xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ");
        sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
        sb.Append("xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
        foreach (var name in names)
        {
            sb.Append("<w:font w:name=\"").Append(System.Security.SecurityElement.Escape(name)).Append("\">");
            sb.Append(KnownFontInner.TryGetValue(name, out var inner) ? inner : GenericInner(name));
            sb.Append("</w:font>");
        }
        sb.Append("</w:fonts>");

        var part = main.AddNewPart<FontTablePart>("rIdFontTableBackfill");
        using var writer = new StreamWriter(part.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false));
        writer.Write(sb.ToString());
    }

    private static void BackfillWebSettings(MainDocumentPart main)
    {
        var part = main.AddNewPart<WebSettingsPart>("rIdWebSettingsBackfill");
        using var writer = new StreamWriter(part.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false));
        writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:webSettings xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
    }

    private static IEnumerable<XElement> EnumerateFontSources(MainDocumentPart main)
    {
        yield return main.GetXDocument().Root!;
        if (main.StyleDefinitionsPart is { } styles)
            yield return styles.GetXDocument().Root!;
    }
}
